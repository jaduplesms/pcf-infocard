"""Discover prerequisites and safely seed eight fictional rows; dry-run by default."""
import argparse
import hashlib
import json
import time
import uuid
from pathlib import Path
from urllib.parse import urlparse

from sample_data import MARKER, SOLUTION, TABLES, fixture_ids, identifier, records

ROOT = Path(__file__).resolve().parent


def environment_url(value):
    parsed = urlparse(value)
    if (parsed.scheme != "https" or not parsed.hostname or parsed.username or parsed.password
            or parsed.port or parsed.path not in ("", "/") or parsed.query or parsed.fragment
            or "YOUR-" in value.upper()):
        raise ValueError("environment_url must be an HTTPS Dataverse origin, without path or credentials.")
    return value.rstrip("/").lower()


class Client:
    def __init__(self, config):
        import msal
        import requests

        self.environment = environment_url(config["environment_url"])
        tenant = str(uuid.UUID(config["tenant_id"]))
        client_id = str(uuid.UUID(config["client_id"]))
        key = hashlib.sha256((self.environment + tenant + client_id).encode()).hexdigest()[:20]
        cache_path = ROOT / "output" / key / ".msal-cache.json"
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache = msal.SerializableTokenCache()
        if cache_path.exists():
            cache.deserialize(cache_path.read_text(encoding="utf-8"))
        app = msal.PublicClientApplication(client_id, authority="https://login.microsoftonline.com/" + tenant,
                                           token_cache=cache)
        accounts = app.get_accounts()
        if len(accounts) > 1:
            raise RuntimeError("Multiple cached identities: remove this sample's cache to select explicitly.")
        scopes = [self.environment + "/.default"]
        token = app.acquire_token_silent(scopes, account=accounts[0]) if accounts else None
        if not token or "access_token" not in token:
            print("Opening Microsoft sign-in for the explicitly configured environment.", flush=True)
            token = app.acquire_token_interactive(scopes=scopes, prompt="select_account", timeout=300)
        if "access_token" not in token:
            raise RuntimeError(token.get("error_description", "Authentication failed."))
        cache_path.write_text(cache.serialize(), encoding="utf-8")
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": "Bearer " + token["access_token"], "Accept": "application/json",
            "Content-Type": "application/json", "OData-MaxVersion": "4.0", "OData-Version": "4.0",
        })
        self.api = self.environment + "/api/data/v9.2/"

    def request(self, method, path, body=None, params=None, headers=None, allowed=()):
        if path.startswith("/") or "://" in path:
            raise ValueError("API paths must remain relative to the configured environment.")
        for attempt in range(4):
            response = self.session.request(method, self.api + path, json=body, params=params,
                                            headers=headers, timeout=180, allow_redirects=False)
            if response.status_code == 429 and attempt < 3:
                time.sleep(float(response.headers.get("Retry-After", 10)))
                continue
            if response.status_code not in (200, 201, 204, *allowed):
                raise RuntimeError(f"{method} {path}: HTTP {response.status_code}: {response.text[:1500]}")
            return response
        raise RuntimeError("Dataverse retry limit reached.")

    def get(self, path, **params):
        return self.request("GET", path, params=params).json()


def discover(client):
    metadata = {}
    for table in (*TABLES, "organization"):
        metadata[table] = client.get(f"EntityDefinitions(LogicalName='{table}')", **{
            "$select": "EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute,OwnershipType",
            "$expand": "Attributes($select=LogicalName,IsValidForCreate,RequiredLevel),"
                       "ManyToOneRelationships($select=ReferencingAttribute,ReferencedEntity,"
                       "ReferencingEntityNavigationPropertyName)",
        })
    return metadata


def bind_payload(table, payload, metadata):
    meta = metadata[table]
    attributes = {attribute["LogicalName"]: attribute for attribute in meta["Attributes"]}
    result = {}
    for field, value in payload.items():
        if field not in attributes or not attributes[field].get("IsValidForCreate"):
            raise ValueError(f"Missing or non-createable prerequisite column: {table}.{field}")
        if isinstance(value, dict):
            relations = [r for r in meta["ManyToOneRelationships"]
                         if r["ReferencingAttribute"] == field and r["ReferencedEntity"] == value["table"]]
            if len(relations) != 1:
                raise ValueError(f"Ambiguous/missing lookup metadata: {table}.{field} -> {value['table']}")
            navigation = relations[0]["ReferencingEntityNavigationPropertyName"]
            result[navigation + "@odata.bind"] = f"/{metadata[value['table']]['EntitySetName']}({value['id']})"
        else:
            result[field] = value
    return result


def existing(client, table, record_id, metadata, payload):
    fields = {metadata[table]["PrimaryIdAttribute"]}
    fields.update("_" + key + "_value" if isinstance(value, dict) else key for key, value in payload.items())
    response = client.request("GET", f"{metadata[table]['EntitySetName']}({record_id})",
                              params={"$select": ",".join(sorted(fields))}, allowed=(404,))
    if response.status_code == 404:
        error = response.json().get("error", {})
        if error.get("code", "").lower() != "0x80040217":
            raise RuntimeError("Unexpected 404; refusing to treat an API/schema error as a missing row.")
        return None
    return response.json()


def validate_record(table, actual, payload):
    # Existing dates/images may legitimately normalize or age; never reset them on a rerun.
    names = {"account": "name", "contact": "lastname", "msdyn_workorder": "msdyn_name",
             "pricelevel": "name", "msdyn_workordertype": "msdyn_name"}
    name = names[table]
    if actual.get(name) != payload[name]:
        raise RuntimeError(f"Sample ID collision or edited identity in {table}; no overwrite performed.")
    for field, expected in payload.items():
        if isinstance(expected, dict):
            if actual.get("_" + field + "_value", "").lower() != expected["id"].lower():
                raise RuntimeError(f"Unexpected lookup in {table}.{field}; no overwrite performed.")
    for field in ("msdyn_systemstatus", "msdyn_incidentrequired", "msdyn_serviceaccountrequired", "msdyn_module"):
        if field in payload and actual.get(field) != payload[field]:
            raise RuntimeError(f"Unexpected prerequisite/lifecycle value in {table}.{field}.")
    marker_field = {"account": "description", "contact": "description", "pricelevel": "description",
                    "msdyn_workorder": "msdyn_workordersummary"}.get(table)
    if marker_field and MARKER not in (actual.get(marker_field) or ""):
        raise RuntimeError(f"Missing fictional marker in {table}; refusing to reuse the row.")


def seed(client, apply=False):
    identity = client.get("WhoAmI()")
    user_id = str(uuid.UUID(identity["UserId"]))
    metadata = discover(client)
    organizations = client.get(metadata["organization"]["EntitySetName"], **{
        "$select": "_basecurrencyid_value", "$top": 2,
    })["value"]
    if len(organizations) != 1 or not organizations[0].get("_basecurrencyid_value"):
        raise RuntimeError("Cannot uniquely resolve the target organization's base currency.")
    currency_id = str(uuid.UUID(organizations[0]["_basecurrencyid_value"]))
    client.get(f"{metadata['transactioncurrency']['EntitySetName']}({currency_id})",
               **{"$select": "transactioncurrencyid"})
    options = client.get(
        "EntityDefinitions(LogicalName='pricelevel')/Attributes(LogicalName='msdyn_module')/"
        "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
        **{"$expand": "OptionSet($select=Options)"},
    )["OptionSet"]["Options"]
    sales = [option["Value"] for option in options if any(
        label["LanguageCode"] == 1033 and label["Label"] == "Sales"
        for label in option["Label"]["LocalizedLabels"])]
    if len(sales) != 1:
        raise RuntimeError("Cannot uniquely resolve the Sales price-list context.")
    plan = records(currency_id, sales[0])
    prepared = []
    for table, record_id, payload in plan:
        body = bind_payload(table, payload, metadata)
        if metadata[table]["OwnershipType"] == "UserOwned":
            body.update(bind_payload(table, {"ownerid": {"table": "systemuser", "id": user_id}}, metadata))
        required = [a["LogicalName"] for a in metadata[table]["Attributes"]
                    if a.get("IsValidForCreate") and a["RequiredLevel"]["Value"] in ("SystemRequired", "ApplicationRequired")]
        provided = set(payload) | {metadata[table]["PrimaryIdAttribute"], "ownerid", "statecode", "statuscode"}
        missing = set(required) - provided
        if missing:
            raise RuntimeError(f"Additional required columns in {table}: {sorted(missing)}. No writes started.")
        current = existing(client, table, record_id, metadata, payload)
        if current:
            validate_record(table, current, payload)
        prepared.append((table, record_id, payload, body, current))
    print(f"Target: {client.environment}; fictional plan: 2 accounts, 2 contacts, 2 unscheduled work orders, "
          "1 price list, 1 work-order type.", flush=True)
    for table, record_id, payload, body, current in prepared:
        if current:
            print(f"SKIP {table} {record_id} (verified existing sample)", flush=True)
            continue
        if not apply:
            print(f"PLAN create {table} {record_id}", flush=True)
            continue
        response = client.request("PATCH", f"{metadata[table]['EntitySetName']}({record_id})", body,
                                  headers={"If-None-Match": "*"}, allowed=(412,))
        actual = existing(client, table, record_id, metadata, payload)
        if actual is None:
            raise RuntimeError(f"Created {table} row is not readable after write.")
        validate_record(table, actual, payload)
        for field, value in payload.items():
            if not isinstance(value, dict) and field != "entityimage" and response.status_code != 412:
                stored = actual.get(field)
                if field in ("birthdate", "anniversary"):
                    stored = stored[:10] if stored else stored
                if stored != value:
                    raise RuntimeError(f"Created value did not persist: {table}.{field}")
        if "entityimage" in payload and not actual.get("entityimage"):
            raise RuntimeError(f"Populated {table} portrait did not persist.")
        print(f"{'SKIP' if response.status_code == 412 else 'CREATE'} {table} {record_id} (read-back verified)", flush=True)
    if not apply:
        print("Dry-run complete. Use --apply only in a disposable Field Service development environment.")
        return None
    # Fill only empty primary-contact links on our marked fixtures, with concurrency guards.
    for kind in ("populated", "sparse"):
        account_id = identifier(f"account:{kind}")
        contact_id = identifier(f"contact:{kind}")
        path = f"{metadata['account']['EntitySetName']}({account_id})"
        account = client.get(path, **{"$select": "description,_primarycontactid_value,_ownerid_value"})
        if MARKER not in (account.get("description") or "") or account.get("_ownerid_value") != user_id:
            raise RuntimeError("Fixture ownership/marker changed; refusing to repair links or build offline scope.")
        linked = account.get("_primarycontactid_value")
        if linked and linked != contact_id:
            raise RuntimeError("Sample primary contact was changed; no overwrite performed.")
        if not linked:
            body = bind_payload("account", {"primarycontactid": {"table": "contact", "id": contact_id}}, metadata)
            client.request("PATCH", path, body, headers={"If-Match": account["@odata.etag"]})
        if client.get(path, **{"$select": "_primarycontactid_value"})["_primarycontactid_value"] != contact_id:
            raise RuntimeError("Sample primary-contact link did not persist.")
    for table, record_id, payload, body, current in prepared:
        actual = existing(client, table, record_id, metadata, payload)
        validate_record(table, actual, payload)
        if metadata[table]["OwnershipType"] == "UserOwned":
            owner = client.get(f"{metadata[table]['EntitySetName']}({record_id})", **{"$select": "_ownerid_value"})
            if owner.get("_ownerid_value") != user_id:
                raise RuntimeError("Sample rows have another owner; this user's offline profile would be incomplete.")
    scope = {table: fixture_ids(table) for table in ("account", "contact", "msdyn_workorder")}
    scope.update(pricelevel=[identifier("pricelevel")], msdyn_workordertype=[identifier("msdyn_workordertype")])
    return {"solution": SOLUTION, "verified": True, "environment_url": client.environment,
            "currency_id": currency_id, "user_id": user_id, "records": scope}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=ROOT / "config.json")
    parser.add_argument("--apply", action="store_true", help="Create only this sample's eight fictional rows.")
    args = parser.parse_args()
    config = json.loads(args.config.read_text(encoding="utf-8"))
    client = Client(config)
    output = ROOT / "output" / hashlib.sha256(client.environment.encode()).hexdigest()[:20] / "seed-manifest.json"
    if args.apply and output.exists():
        output.unlink()
    manifest = seed(client, args.apply)
    if manifest:
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(manifest, indent=2), encoding="utf-8")
        print(f"Verified offline-build manifest: {output}")
