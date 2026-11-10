"""Pack the independent sample app; never connects to Dataverse."""
import argparse
import json
import shutil
import subprocess
import tempfile
import uuid
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from sample_data import APP, SOLUTION, TABLES, fixture_ids, identifier

ROOT = Path(__file__).resolve().parent


def text(parent, name, value):
    ET.SubElement(parent, name).text = str(value)


def write_xml(path, root):
    path.parent.mkdir(parents=True, exist_ok=True)
    ET.indent(root, space="  ")
    ET.ElementTree(root).write(path, encoding="utf-8", xml_declaration=True)


def fetch(table, ids):
    root = ET.Element("fetch")
    entity = ET.SubElement(root, "entity", {"name": table})
    key = TABLES[table][1]
    ET.SubElement(entity, "attribute", {"name": key})
    conditions = ET.SubElement(entity, "filter", {"type": "or"})
    for record_id in ids:
        ET.SubElement(conditions, "condition", {
            "attribute": key, "operator": "eq", "value": str(uuid.UUID(record_id)),
        })
    return ET.tostring(root, encoding="unicode")


def localized(parent, container, element, value):
    ET.SubElement(ET.SubElement(parent, container), element, {
        "description": value, "languagecode": "1033",
    })


def scaffold(source, offline):
    """Canonical component shapes; only the dedicated form bindings are stored."""
    version = "1.0.0.1" if offline else "1.0.0.0"
    solution = ET.Element("ImportExportXml", {
        "version": "9.2", "SolutionPackageVersion": "9.2", "languagecode": "1033",
        "generatedBy": "CrmLive",
    })
    manifest = ET.SubElement(solution, "SolutionManifest")
    text(manifest, "UniqueName", SOLUTION)
    localized(manifest, "LocalizedNames", "LocalizedName", "InfoCard Test App (SAMPLE ONLY)")
    localized(manifest, "Descriptions", "Description", "Optional fictional test app; not the shipping control.")
    text(manifest, "Version", version)
    text(manifest, "Managed", 0)
    publisher = ET.SubElement(manifest, "Publisher")
    text(publisher, "UniqueName", "InfoCardTestAppPublisher")
    localized(publisher, "LocalizedNames", "LocalizedName", "InfoCard Test Samples")
    localized(publisher, "Descriptions", "Description", "Development-only fictional examples.")
    text(publisher, "CustomizationPrefix", "ictest")
    text(publisher, "CustomizationOptionValuePrefix", 89124)
    roots = ET.SubElement(manifest, "RootComponents")
    ET.SubElement(manifest, "MissingDependencies")
    app = ET.Element("AppModule")
    for key, value in (("UniqueName", APP), ("IntroducedVersion", version),
                       ("WebResourceId", "953b9fac-1e5e-e611-80d6-00155ded156f"),
                       ("statecode", 0), ("statuscode", 1), ("FormFactor", 1),
                       ("ClientType", 4), ("NavigationType", 0)):
        text(app, key, value)
    components = ET.SubElement(app, "AppModuleComponents")
    sitemap = ET.Element("AppModuleSiteMap")
    text(sitemap, "SiteMapUniqueName", APP)
    for key in ("EnableCollapsibleGroups", "ShowHome", "ShowPinned", "ShowRecents"):
        text(sitemap, key, "False" if key == "EnableCollapsibleGroups" else "True")
    site = ET.SubElement(sitemap, "SiteMap")
    area = ET.SubElement(site, "Area", {"Id": "ictest_area", "ShowGroups": "true"})
    ET.SubElement(ET.SubElement(area, "Titles"), "Title", {"LCID": "1033", "Title": "InfoCard Test App"})
    group = ET.SubElement(area, "Group", {"Id": "ictest_fixtures"})
    ET.SubElement(ET.SubElement(group, "Titles"), "Title", {
        "LCID": "1033", "Title": "Fictional test records",
    })
    for table in ("account", "contact", "msdyn_workorder"):
        schema, primary, name = TABLES[table]
        label = {"account": "Account", "contact": "Contact", "msdyn_workorder": "Work Order"}[table]
        directory = source / "Entities" / schema
        entity = ET.Element("Entity")
        ET.SubElement(entity, "Name", {"LocalizedName": label, "OriginalName": ""}).text = schema
        ET.SubElement(ET.SubElement(ET.SubElement(entity, "EntityInfo"), "entity", {
            "Name": schema, "unmodified": "1",
        }), "attributes")
        ET.SubElement(entity, "FormXml")
        ET.SubElement(entity, "SavedQueries")
        ET.SubElement(entity, "RibbonDiffXml")
        write_xml(directory / "Entity.xml", entity)
        ET.SubElement(roots, "RootComponent", {"type": "1", "schemaName": table, "behavior": "1"})
        ET.SubElement(components, "AppModuleComponent", {"type": "1", "schemaName": table})
        for kind, component_type in (("form", "60"), ("view", "26")):
            component_id = "{" + identifier(f"{kind}:{table}") + "}"
            ET.SubElement(roots, "RootComponent", {"type": component_type, "id": component_id, "behavior": "0"})
            ET.SubElement(components, "AppModuleComponent", {"type": component_type, "id": component_id})
        savedqueries = ET.Element("savedqueries")
        view = ET.SubElement(savedqueries, "savedquery")
        for key, value in (("IsCustomizable", 1), ("CanBeDeleted", 1), ("isquickfindquery", 0),
                           ("isprivate", 0), ("isdefault", 0),
                           ("savedqueryid", "{" + identifier(f"view:{table}") + "}")):
            text(view, key, value)
        grid = ET.SubElement(ET.SubElement(view, "layoutxml"), "grid", {
            "name": "resultset", "jump": name, "select": "1", "icon": "1", "preview": "1",
        })
        row = ET.SubElement(grid, "row", {"name": "result", "id": primary})
        for column, width in ((name, "350"), ("createdon", "150")):
            ET.SubElement(row, "cell", {"name": column, "width": width})
        text(view, "querytype", 0)
        query = ET.fromstring(fetch(table, fixture_ids(table)))
        query.set("mapping", "logical")
        query_entity = query.find("entity")
        for column in (name, "createdon"):
            ET.SubElement(query_entity, "attribute", {"name": column})
        ET.SubElement(query_entity, "order", {"attribute": name})
        ET.SubElement(view, "fetchxml").append(query)
        text(view, "IntroducedVersion", version)
        localized(view, "LocalizedNames", "LocalizedName", "InfoCard Sample - " + label + "s")
        localized(view, "Descriptions", "Description", "Only the two fictional sample records.")
        write_xml(directory / "SavedQueries" / (identifier(f"view:{table}") + ".xml"), savedqueries)
        subarea = ET.SubElement(group, "SubArea", {"Id": "ictest_" + table, "Entity": table})
        ET.SubElement(ET.SubElement(subarea, "Titles"), "Title", {"LCID": "1033", "Title": label + "s"})
    for component_type in ("62", "80"):
        ET.SubElement(roots, "RootComponent", {"type": component_type, "schemaName": APP, "behavior": "0"})
    ET.SubElement(components, "AppModuleComponent", {"type": "62", "schemaName": APP})
    localized(app, "LocalizedNames", "LocalizedName", "InfoCard Test App")
    localized(app, "Descriptions", "Description", "Fictional examples for web and Field Service Mobile.")
    settings = ET.SubElement(app, "appsettings")
    for key, value in (("AppChannel", "1"), ("ChannelAvailability", "FieldServiceMobile")):
        setting = ET.SubElement(settings, "appsetting", {"settingdefinitionid.uniquename": key})
        text(setting, "iscustomizable", 1)
        text(setting, "value", value)
    localized(sitemap, "LocalizedNames", "LocalizedName", "InfoCard Test App")
    write_xml(source / "AppModules" / APP / "AppModule.xml", app)
    write_xml(source / "AppModuleSiteMaps" / APP / "AppModuleSiteMap.xml", sitemap)
    write_xml(source / "Other" / "Solution.xml", solution)
    customization = ET.Element("ImportExportXml")
    for container in ("Entities", "Roles", "Workflows", "FieldSecurityProfiles", "Templates",
                      "EntityImageConfigs", "EntityMaps", "EntityRelationships", "OrganizationSettings",
                      "optionsets", "MobileOfflineProfiles", "CustomControls", "AppModuleSiteMaps",
                      "AppModules", "EntityDataProviders"):
        ET.SubElement(customization, container)
    text(ET.SubElement(customization, "Languages"), "Language", 1033)
    write_xml(source / "Other" / "Customizations.xml", customization)


def add_offline(source, manifest):
    if manifest.get("solution") != SOLUTION or manifest.get("verified") is not True:
        raise ValueError("Offline build requires the successful seeder's verified manifest.")
    currency_id = str(uuid.UUID(manifest["currency_id"]))
    user_id = str(uuid.UUID(manifest["user_id"]))
    scope = {table: fixture_ids(table) for table in ("account", "contact", "msdyn_workorder")}
    scope.update(transactioncurrency=[currency_id], systemuser=[user_id],
                 pricelevel=[identifier("pricelevel")],
                 msdyn_workordertype=[identifier("msdyn_workordertype")])
    expected = {table: ids for table, ids in scope.items() if table not in ("systemuser", "transactioncurrency")}
    if manifest.get("records") != expected:
        raise ValueError("Seeder manifest record IDs do not match this sample.")
    # A changed target/user/currency gets fresh profile identity, avoiding stale imported filters.
    profile_id = identifier(f"profile:{manifest['environment_url']}:{currency_id}:{user_id}")
    name = "InfoCard Sample Offline"
    root = ET.Element("MobileOfflineProfile", {"MobileOfflineProfileId": profile_id, "Name": name})
    items = ET.SubElement(root, "MobileOfflineProfileItems")
    for table, ids in scope.items():
        schema = TABLES[table][0]
        item_id = identifier(f"item:{profile_id}:{table}")
        item = ET.SubElement(items, "MobileOfflineProfileItem", {
            "MobileOfflineProfileItemId": item_id, "Name": schema,
        })
        ET.SubElement(item, "MobileOfflineProfileItemAssociations")
        for key, value in (
            ("MobileOfflineProfileItemId", "{" + item_id + "}"),
            ("RegardingObjectId", "{" + profile_id + "}"), ("Name", schema),
            ("RecordDistributionCriteria", 3), ("RecordsOwnedByMe", 0),
            ("RecordsOwnedByMyTeam", 0), ("RecordsOwnedByMyBusinessUnit", 0),
            ("ProfileItemEntityFilter", fetch(table, ids)), ("EntitySchemaName", schema),
        ):
            text(item, key, value)
    text(root, "MobileOfflineProfileId", "{" + profile_id + "}")
    text(root, "Name", name)
    write_xml(source / "MobileOfflineProfiles" / (name + ".xml"), root)
    app_path = source / "AppModules" / APP / "AppModule.xml"
    app = ET.parse(app_path).getroot()
    ET.SubElement(app.find("AppModuleComponents"), "AppModuleComponent", {
        "type": "161", "id": "{" + profile_id + "}",
    })
    settings = app.find("appsettings")
    for key, value in (("EnableRoleBasedOfflineProfileAccess", "true"), ("ServerMode", "false")):
        setting = ET.SubElement(settings, "appsetting", {"settingdefinitionid.uniquename": key})
        text(setting, "iscustomizable", 1)
        text(setting, "value", value)
    write_xml(app_path, app)
    solution_path = source / "Other" / "Solution.xml"
    solution = ET.parse(solution_path).getroot()
    ET.SubElement(solution.find("SolutionManifest/RootComponents"), "RootComponent", {
        "type": "161", "id": "{" + profile_id + "}", "behavior": "0",
    })
    write_xml(solution_path, solution)
    return profile_id


def verify_package(path, offline=False):
    with zipfile.ZipFile(path) as archive:
        solution = ET.fromstring(archive.read("solution.xml"))
        customization = ET.fromstring(archive.read("customizations.xml"))
        if solution.findtext("SolutionManifest/UniqueName") != SOLUTION:
            raise ValueError("Incorrect sample solution identity.")
        if solution.findtext("SolutionManifest/Managed") != "0":
            raise ValueError("The development sample must be unmanaged.")
        forms = customization.findall(".//systemform")
        views = customization.findall(".//savedquery")
        if len(forms) != 3 or len(views) != 3:
            raise ValueError("Package must include all three dedicated forms and views.")
        for table in ("account", "contact", "msdyn_workorder"):
            if not any(f.findtext("formid", "").strip("{}") == identifier(f"form:{table}") for f in forms):
                raise ValueError(f"Missing dedicated {table} form.")
            if not any(v.findtext("savedqueryid", "").strip("{}") == identifier(f"view:{table}") for v in views):
                raise ValueError(f"Missing dedicated {table} view.")
        for form in forms:
            for description in form.findall(".//controlDescription"):
                if description.find("customControl[@name='smp_Sample.InfoCard']") is None:
                    continue
                control_id = description.get("forControl", "").strip("{}").lower()
                controls = [control for control in form.findall(".//control")
                            if control.get("uniqueid", "").strip("{}").lower() == control_id]
                if len(controls) != 1 or controls[0].get("disabled", "").lower() in ("true", "1"):
                    raise ValueError("InfoCard placements must exist and have the read-only flag cleared.")
        apps = customization.findall(".//AppModule")
        if len(apps) != 1 or apps[0].findtext("UniqueName") != APP or apps[0].findtext("NavigationType") != "0":
            raise ValueError("Expected one single-session sample app.")
        if customization.findall(".//formLibraries/Library") or customization.findall(".//events/event"):
            raise ValueError("Sample unexpectedly includes inherited form scripts.")
        if any(name.startswith("Controls/") for name in archive.namelist()):
            raise ValueError("The sample must reference, not bundle, the shipping control.")
        profile_items = customization.findall(".//MobileOfflineProfileItem")
        if len(profile_items) != (7 if offline else 0):
            raise ValueError("Offline profile composition is incorrect.")
        for item in profile_items:
            if item.findtext("RecordDistributionCriteria") != "3":
                raise ValueError("Offline profile scope was broadened.")
            if not ET.fromstring(item.findtext("ProfileItemEntityFilter")).findall(".//condition"):
                raise ValueError("An offline table lacks an exact-record filter.")


def build(output, manifest=None):
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    pac = shutil.which("pac")
    if not pac:
        raise RuntimeError("Power Platform CLI (pac) is required to pack the sample.")
    with tempfile.TemporaryDirectory(prefix="infocard-sample-") as stage:
        source = Path(stage) / "source"
        shutil.copytree(ROOT / "solution", source)
        scaffold(source, manifest is not None)
        profile_id = add_offline(source, manifest) if manifest else None
        packed = Path(stage) / "InfoCardTestApp.zip"
        subprocess.run([pac, "solution", "pack", "--folder", str(source), "--zipfile", str(packed),
                        "--packagetype", "Unmanaged", "--errorlevel", "Verbose"], check=True)
        if not packed.is_file():
            raise RuntimeError("PAC did not produce the sample package; review the pack output.")
        verify_package(packed, offline=manifest is not None)
        shutil.copyfile(packed, output)
    print(f"Verified package: {output}")
    if profile_id:
        print(f"Target-specific offline profile: {profile_id}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out", type=Path, default=ROOT / "output" / "InfoCardTestApp.zip")
    parser.add_argument("--offline-manifest", type=Path, help="Verified output from seed.py; target-specific ZIP.")
    args = parser.parse_args()
    manifest = json.loads(args.offline_manifest.read_text(encoding="utf-8")) if args.offline_manifest else None
    build(args.out, manifest)
