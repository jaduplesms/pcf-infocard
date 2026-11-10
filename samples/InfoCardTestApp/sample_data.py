"""Fictional sample identity and payloads shared by the package and seeder."""
import base64
import uuid
from datetime import date, timedelta
from pathlib import Path

NAMESPACE = uuid.UUID("f668278c-92bb-43e8-8a3f-c68a79d4c5e6")
SOLUTION = "InfoCardTestApp"
APP = "ictest_InfoCardTestApp"
MARKER = "Harbourview"
TABLES = {
    "account": ("Account", "accountid", "name"),
    "contact": ("Contact", "contactid", "fullname"),
    "msdyn_workorder": ("msdyn_workorder", "msdyn_workorderid", "msdyn_name"),
    "pricelevel": ("PriceLevel", "pricelevelid", "name"),
    "msdyn_workordertype": ("msdyn_workordertype", "msdyn_workordertypeid", "msdyn_name"),
    "transactioncurrency": ("TransactionCurrency", "transactioncurrencyid", "currencyname"),
    "systemuser": ("SystemUser", "systemuserid", "fullname"),
}


def identifier(key):
    return str(uuid.uuid5(NAMESPACE, key))


def fixture_ids(table):
    return [identifier(f"{table}:{kind}") for kind in ("populated", "sparse")]


def portrait():
    """Synthetic portrait stored locally so the demo has no image-CDN dependency."""
    path = Path(__file__).parent / "assets" / "maya-bennett.jpg"
    return base64.b64encode(path.read_bytes()).decode("ascii")


def account_image():
    """Synthetic building photograph for the business account."""
    path = Path(__file__).parent / "assets" / "harbourview-business-centre.jpg"
    return base64.b64encode(path.read_bytes()).decode("ascii")


def records(currency_id, sales_module, today=None):
    """Logical lookup descriptors are resolved through target metadata."""
    today = today or date.today()
    currency = {"table": "transactioncurrency", "id": currency_id}
    result = [
        ("pricelevel", identifier("pricelevel"), {
            "name": "Commercial Maintenance Rates",
            "description": "Harbourview commercial service rates for planned building maintenance.",
            "msdyn_module": sales_module,
            "transactioncurrencyid": currency,
        }),
        ("msdyn_workordertype", identifier("msdyn_workordertype"), {
            "msdyn_name": "Planned Maintenance",
            "msdyn_incidentrequired": False, "msdyn_serviceaccountrequired": True,
            "msdyn_taxable": False,
        }),
    ]
    address = {
        "address1_line1": "18 Willow Quay", "address1_city": "Bristol",
        "address1_postalcode": "BS1 4QG", "address1_country": "United Kingdom",
    }
    image = portrait()
    for kind in ("populated", "sparse"):
        account = {"name": "Harbourview Business Centre" if kind == "populated" else "Harbourview Annex",
                   "description": "Harbourview Annex is awaiting its initial facilities survey.",
                   "transactioncurrencyid": currency}
        if kind == "populated":
            account.update(address, telephone1="+44 1632 960101", telephone2="+44 1632 960102",
                           emailaddress1="facilities@harbourview.example.com",
                           websiteurl="https://example.com/harbourview", entityimage=account_image(), revenue=2450000,
                           numberofemployees=48,
                           description="Harbourview Business Centre is a six-storey office campus. "
                           "Planned maintenance covers air handling, pumps and building controls.")
        result.append(("account", identifier(f"account:{kind}"), account))
    for kind in ("populated", "sparse"):
        contact = {
            "firstname": "Maya" if kind == "populated" else "Alex",
            "lastname": "Bennett" if kind == "populated" else "Grant",
            "description": "Harbourview facilities contact; preferred communication details are being confirmed.",
            "parentcustomerid": {"table": "account", "id": identifier(f"account:{kind}")},
        }
        if kind == "populated":
            contact.update(address, jobtitle="Facilities Operations Coordinator",
                           mobilephone="+44 7700 900101", telephone1="+44 1632 960103",
                           emailaddress1="maya.bennett@harbourview.example.com",
                           websiteurl="https://example.com/harbourview/team/maya-bennett", entityimage=image,
                           description="Harbourview's facilities coordinator and first point of contact "
                           "for planned visits. Available weekdays, 08:00-17:00; call ahead for plant-room access.")
        result.append(("contact", identifier(f"contact:{kind}"), contact))
    for kind in ("populated", "sparse"):
        account = {"table": "account", "id": identifier(f"account:{kind}")}
        order = {
            "msdyn_name": "WO-10428" if kind == "populated" else "WO-10429",
            "msdyn_serviceaccount": account, "msdyn_billingaccount": account,
            "msdyn_pricelist": {"table": "pricelevel", "id": identifier("pricelevel")},
            "msdyn_workordertype": {"table": "msdyn_workordertype", "id": identifier("msdyn_workordertype")},
            "transactioncurrencyid": currency, "msdyn_systemstatus": 690970000,
            "msdyn_taxable": False,
            "msdyn_workordersummary": "Initial facilities survey at Harbourview Annex.",
        }
        if kind == "populated":
            day = (today + timedelta(days=7)).isoformat()
            order.update(msdyn_phoneNumber="+44 1632 960104",
                         msdyn_address1="18 Willow Quay", msdyn_city="Bristol",
                         msdyn_postalcode="BS1 4QG", msdyn_country="United Kingdom",
                         msdyn_datewindowstart=day + "T09:00:00Z",
                         msdyn_datewindowend=day + "T12:00:00Z",
                         msdyn_workordersummary="Quarterly air-handling service at Harbourview Business Centre.",
                         msdyn_instructions="Report to reception and ask for Maya Bennett. "
                         "Plant-room access is via the east service lift. "
                         "Record filter condition and operating readings in the service report.")
        result.append(("msdyn_workorder", identifier(f"msdyn_workorder:{kind}"), order))
    return result
