#!/usr/bin/env npx ts-node
/**
 * InfoCard Configuration Tool
 *
 * Extract and apply InfoCard PCF control configurations between Dataverse forms.
 *
 * Usage:
 *   npx ts-node tools/infocard-config.ts extract --form <formId> [--output <file.json>]
 *   npx ts-node tools/infocard-config.ts apply --form <formId> --config <file.json> [--field <fieldName>]
 *   npx ts-node tools/infocard-config.ts list-forms --entity <entityName>
 *
 * Requires: pac CLI authenticated (pac auth list)
 */

import { execSync } from "child_process";
import * as fs from "fs";
import * as os from "os";

// ────────────────────────────────────────
// Types
// ────────────────────────────────────────

interface InfoCardConfig {
    controlName: string;
    extractedFrom: {
        formId: string;
        formName: string;
        entityName: string;
        fieldName: string;
        extractedAt: string;
    };
    formFactors: {
        phone?: Record<string, ParameterBinding>;
        tablet?: Record<string, ParameterBinding>;
        desktop?: Record<string, ParameterBinding>;
    };
}

interface ParameterBinding {
    type: string;
    value: string;
    static?: boolean;
}

// ────────────────────────────────────────
// PAC CLI helpers
// ────────────────────────────────────────

function pacFetch(fetchXml: string): string {
    const result = execSync(`pac org fetch --xml "${fetchXml.replace(/"/g, '\\"')}"`, {
        encoding: "utf-8",
        timeout: 30000,
    });
    return result;
}

function pacFetchJson(fetchXml: string): string {
    // pac org fetch doesn't support --output json in all versions
    // Return raw text for parsing
    return pacFetch(fetchXml);
}

// ────────────────────────────────────────
// Form XML parsing
// ────────────────────────────────────────

function fetchFormXml(formId: string): { formxml: string; name: string; entityName: string } {
    const fetchXml = `<fetch count='1' no-lock='true'><entity name='systemform'><attribute name='formxml'/><attribute name='name'/><attribute name='objecttypecode'/><filter><condition attribute='formid' operator='eq' value='${formId}'/></filter></entity></fetch>`;
    const raw = pacFetch(fetchXml);

    // The formxml is in the output — extract it
    // pac org fetch returns tabular format, formxml is the content
    const lines = raw.split("\n").filter(l => l.trim());
    // Skip header lines (Connected as..., column headers)
    const dataLine = lines.find(l => l.includes("<tabs>") || l.includes("<form>"));

    if (!dataLine) {
        throw new Error(`Form ${formId} not found or has no form XML`);
    }

    // Extract the form name from the tabular output
    const nameMatch = raw.match(/name\s+objecttypecode\s+formxml/i);
    const formName = "Unknown Form";
    const entityName = "unknown";

    return { formxml: dataLine.trim(), name: formName, entityName };
}

/** Split form XML into individual controlDescription blocks using string split (not regex) */
function splitControlDescriptions(formXml: string): string[] {
    const blocks: string[] = [];
    const parts = formXml.split("<controlDescription");
    for (const part of parts.slice(1)) { // skip everything before first <controlDescription
        const endIdx = part.indexOf("</controlDescription>");
        if (endIdx >= 0) {
            blocks.push("<controlDescription" + part.substring(0, endIdx + "</controlDescription>".length));
        }
    }
    return blocks;
}

function extractInfoCardConfig(formXml: string, formId: string): InfoCardConfig | null {
    // Form XML from pac may be single-line, attributes in any order
    const controlRegex = /<customControl name="smp_Sample\.InfoCard" formFactor="(\d)"><parameters>(.*?)<\/parameters>/g;
    const paramRegex = /<(\w+) type="([^"]*)"(?: static="([^"]*)")?>(.*?)<\/\1>/g;

    const factorMap: Record<string, string> = { "0": "phone", "2": "tablet", "1": "desktop" };
    const formFactors: Record<string, Record<string, ParameterBinding>> = {};
    let fieldName = "";

    // Find the controlDescription block that contains our InfoCard control
    for (const block of splitControlDescriptions(formXml)) {
        if (block.includes("smp_Sample.InfoCard")) {
            const fnMatch = block.match(/<datafieldname>([^<]+)<\/datafieldname>/);
            if (fnMatch) fieldName = fnMatch[1];
            break;
        }
    }

    let match;
    while ((match = controlRegex.exec(formXml)) !== null) {
        const formFactor = factorMap[match[1]] ?? `factor${match[1]}`;
        const paramsXml = match[2];
        const params: Record<string, ParameterBinding> = {};

        let paramMatch;
        while ((paramMatch = paramRegex.exec(paramsXml)) !== null) {
            params[paramMatch[1]] = {
                type: paramMatch[2],
                value: paramMatch[4],
                static: paramMatch[3] === "true" ? true : undefined,
            };
        }

        formFactors[formFactor] = params;
    }

    if (Object.keys(formFactors).length === 0) return null;

    return {
        controlName: "smp_Sample.InfoCard",
        extractedFrom: {
            formId,
            formName: "",
            entityName: "",
            fieldName,
            extractedAt: new Date().toISOString(),
        },
        formFactors: {
            phone: formFactors.phone,
            tablet: formFactors.tablet,
            desktop: formFactors.desktop,
        },
    };
}

/** Build inline XML for one customControl element (no newlines — matches Dataverse format) */
function buildControlXml(params: Record<string, ParameterBinding>, formFactor: number): string {
    const paramParts = Object.entries(params).map(([name, binding]) => {
        const staticAttr = binding.static ? ` static="true"` : "";
        return `<${name} type="${binding.type}"${staticAttr}>${binding.value}</${name}>`;
    });
    return `<customControl name="smp_Sample.InfoCard" formFactor="${formFactor}"><parameters>${paramParts.join("")}</parameters></customControl>`;
}

/** Build a full controlDescription block with all three form factors */
function buildControlDescriptionXml(config: InfoCardConfig, fieldName: string): string {
    const defaultCtrl = `<customControl id="{270BD3DB-D9AF-4782-9025-509E298DEC0A}"><parameters><datafieldname>${fieldName}</datafieldname></parameters></customControl>`;
    const factors: string[] = [defaultCtrl];
    const factorMap: Record<string, number> = { phone: 0, tablet: 2, desktop: 1 };

    for (const [factorName, factorNum] of Object.entries(factorMap)) {
        const params = config.formFactors[factorName as keyof typeof config.formFactors];
        if (params) {
            factors.push(buildControlXml(params, factorNum));
        }
    }

    return `<controlDescription forControl="{${generateGuid()}}">` + factors.join("") + `</controlDescription>`;
}

function generateGuid(): string {
    return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
        const r = Math.random() * 16 | 0;
        return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
    });
}

// ────────────────────────────────────────
// Commands
// ────────────────────────────────────────

function cmdExtract(formId: string, outputFile?: string): void {
    console.log(`Extracting InfoCard config from form ${formId}...`);

    // Fetch the form XML via WebAPI using pac
    const fetchXml = `<fetch count='1' no-lock='true'><entity name='systemform'><attribute name='formxml'/><attribute name='name'/><attribute name='objecttypecode'/><filter><condition attribute='formid' operator='eq' value='${formId}'/></filter></entity></fetch>`;

    let raw: string;
    try {
        raw = pacFetch(fetchXml);
    } catch (err) {
        console.error("Failed to fetch form. Ensure pac CLI is authenticated (pac auth list)");
        process.exit(1);
    }

    // pac org fetch returns formxml on a data line — find the line containing <form
    const formLine = raw.split("\n").find(l => l.includes("<form"));
    if (!formLine) {
        console.error("No form XML found in response. Check the form ID.");
        process.exit(1);
    }
    const formXml = formLine.trim();

    const config = extractInfoCardConfig(formXml, formId);
    if (!config) {
        console.error("No InfoCard control found on this form");
        process.exit(1);
    }

    const json = JSON.stringify(config, null, 2);

    if (outputFile) {
        fs.writeFileSync(outputFile, json, "utf-8");
        console.log(`Config saved to ${outputFile}`);
    } else {
        console.log(json);
    }

    // Summary
    const factorCount = Object.values(config.formFactors).filter(Boolean).length;
    const paramCount = Object.keys(config.formFactors.phone ?? config.formFactors.desktop ?? {}).length;
    console.log(`\nExtracted: ${factorCount} form factor(s), ${paramCount} parameters each`);
    console.log(`Bound to field: ${config.extractedFrom.fieldName || "(unknown)"}`);
}

function cmdListForms(entityName: string): void {
    console.log(`Listing forms for entity: ${entityName}...`);

    const fetchXml = `<fetch no-lock='true'><entity name='systemform'><attribute name='formid'/><attribute name='name'/><attribute name='type'/><attribute name='description'/><filter><condition attribute='objecttypecode' operator='eq' value='${entityName}'/><condition attribute='type' operator='eq' value='2'/></filter><order attribute='name'/></entity></fetch>`;

    try {
        const raw = pacFetch(fetchXml);
        console.log(raw);
    } catch (err) {
        console.error("Failed to list forms. Ensure pac CLI is authenticated.");
        process.exit(1);
    }
}

// ────────────────────────────────────────
function cmdApply(formId: string, configFile: string, fieldName?: string): void {
    const config: InfoCardConfig = JSON.parse(fs.readFileSync(configFile, "utf-8"));
    const targetField = fieldName ?? config.extractedFrom.fieldName;

    if (!targetField) {
        console.error("No field name specified. Use --field <fieldName> or ensure config has extractedFrom.fieldName");
        process.exit(1);
    }

    console.log(`Applying InfoCard config to form ${formId}, field: ${targetField}`);

    // 1. Fetch current form XML
    const fetchXml = `<fetch count='1' no-lock='true'><entity name='systemform'><attribute name='formxml'/><filter><condition attribute='formid' operator='eq' value='${formId}'/></filter></entity></fetch>`;
    let raw: string;
    try {
        raw = pacFetch(fetchXml);
    } catch (err) {
        console.error("Failed to fetch form. Ensure pac CLI is authenticated.", err);
        process.exit(1);
    }

    const formLine = raw.split("\n").find(l => l.includes("<form"));
    if (!formLine) { console.error("No form XML found"); process.exit(1); }
    let formXml = formLine.trim();

    // 2. Check if InfoCard is already on this form
    const allDescs = splitControlDescriptions(formXml);
    const blocksToRemove: string[] = [];
    for (const desc of allDescs) {
        if (desc.includes(`<datafieldname>${targetField}</datafieldname>`) && desc.includes("smp_Sample.InfoCard")) {
            blocksToRemove.push(desc);
        }
    }

    const newControlDesc = buildControlDescriptionXml(config, targetField);
    const phoneParams = config.formFactors.phone ? Object.keys(config.formFactors.phone).length : 0;
    console.log(`New controlDescription: ${newControlDesc.length} chars, ${phoneParams} phone params`);
    console.log(`Found ${blocksToRemove.length} existing InfoCard block(s) to replace`);

    if (blocksToRemove.length > 0) {
        // Remove ALL existing blocks, then insert the new one in place of the first
        const firstBlock = blocksToRemove[0];
        formXml = formXml.replace(firstBlock, newControlDesc);
        for (let i = 1; i < blocksToRemove.length; i++) {
            formXml = formXml.replace(blocksToRemove[i], "");
        }
        console.log("Replaced existing InfoCard control configuration.");
    } else {
        // Insert new controlDescription — find </controlDescriptions> or create the section
        if (formXml.includes("</controlDescriptions>")) {
            formXml = formXml.replace("</controlDescriptions>", newControlDesc + "</controlDescriptions>");
            console.log("Added InfoCard control to existing controlDescriptions section.");
        } else {
            // No controlDescriptions section — add before </form>
            formXml = formXml.replace("</form>", `<controlDescriptions>${newControlDesc}</controlDescriptions></form>`);
            console.log("Created controlDescriptions section with InfoCard control.");
        }
    }

    // 3. Write updated form XML back to Dataverse
    console.log("Updating form in Dataverse...");
    const escapedXml = formXml.replace(/'/g, "''");

    try {
        // Use pac org fetch to update via FetchXML isn't possible — use a temp file + pac solution approach
        // Instead, write to a temp file and use pac CLI to import
        const tempDir = os.tmpdir();
        const tempXmlFile = `${tempDir}/infocard_form_update.xml`;
        fs.writeFileSync(tempXmlFile, formXml, "utf-8");

        // Use the Dataverse WebAPI via pac to update the systemform record
        // pac doesn't have a direct update command, so we'll use the org update approach
        execSync(
            `pac org fetch --xml "<fetch count='1'><entity name='systemform'><attribute name='formid'/><filter><condition attribute='formid' operator='eq' value='${formId}'/></filter></entity></fetch>"`,
            { encoding: "utf-8", timeout: 15000 },
        );

        // Write the form XML via a PowerShell/curl WebAPI call
        const tokenResult = execSync("pac auth token", { encoding: "utf-8", timeout: 15000 });
        const token = tokenResult.split("\n").filter(l => l.trim() && !l.includes("Connected")).pop()?.trim();

        if (!token) {
            console.error("Could not get auth token. Falling back to file export.");
            const outputPath = `${tempDir}/infocard_updated_form.xml`;
            fs.writeFileSync(outputPath, formXml, "utf-8");
            console.log(`\nUpdated form XML saved to: ${outputPath}`);
            console.log("To apply manually:");
            console.log("1. Import this XML into the systemform record via WebAPI");
            console.log("2. Run: pac solution publish");
            return;
        }

        // Get the org URL from pac org who
        const orgWho = execSync("pac org who", { encoding: "utf-8", timeout: 15000 });
        const orgUrlMatch = orgWho.match(/Org URL:\s*(https?:\/\/[^\s/]+)/i);
        const orgUrl = orgUrlMatch?.[1];

        if (!orgUrl) {
            console.error("Could not determine org URL.");
            process.exit(1);
        }

        // Update via PowerShell Invoke-RestMethod using pac auth token with Dataverse audience
        const jsonBody = JSON.stringify({ formxml: formXml });
        const bodyFile = `${tempDir}/infocard_patch_body.json`;
        fs.writeFileSync(bodyFile, jsonBody, "utf-8");

        // Get token scoped to Dataverse using pac auth token --resource
        const psScript = `
$body = Get-Content -Path '${bodyFile.replace(/'/g, "''")}' -Raw -Encoding UTF8
$tokenRaw = pac auth token --resource '${orgUrl}/' 2>&1 | Select-Object -Last 1
$headers = @{
    'Authorization' = "Bearer $tokenRaw"
    'Content-Type' = 'application/json; charset=utf-8'
    'OData-MaxVersion' = '4.0'
    'OData-Version' = '4.0'
}
try {
    Invoke-RestMethod -Uri '${orgUrl}/api/data/v9.2/systemforms(${formId})' -Method Patch -Headers $headers -Body ([System.Text.Encoding]::UTF8.GetBytes($body))
    Write-Host 'PATCH_OK'
} catch {
    Write-Host "PATCH_FAIL: $_"
}
`;
        const psFile = `${tempDir}/infocard_patch.ps1`;
        fs.writeFileSync(psFile, psScript, "utf-8");

        const result = execSync(`powershell -ExecutionPolicy Bypass -File "${psFile}"`, {
            encoding: "utf-8",
            timeout: 60000,
        });
        if (result.includes("PATCH_OK")) {
            console.log("Form updated successfully.");
        } else {
            console.error("PATCH failed:", result.trim());
            const recoveryPath = `${tempDir}/infocard_recovery_form.xml`;
            fs.writeFileSync(recoveryPath, formXml, "utf-8");
            console.log(`Recovery: updated form XML saved to ${recoveryPath}`);
            process.exit(1);
        }

        // 4. Publish
        console.log("Publishing customizations...");
        execSync("pac solution publish", { encoding: "utf-8", timeout: 120000 });
        console.log("Published.");

        console.log(`\nInfoCard control applied to form ${formId}, field: ${targetField}`);

    } catch (err) {
        console.error("Failed to update form:", err);
        // Save the XML for manual recovery
        const recoveryPath = `${os.tmpdir()}/infocard_recovery_form.xml`;
        fs.writeFileSync(recoveryPath, formXml, "utf-8");
        console.log(`\nRecovery: updated form XML saved to ${recoveryPath}`);
    }
}

// CLI entry point
// ────────────────────────────────────────

const args = process.argv.slice(2);
const command = args[0];

function getArg(name: string): string | undefined {
    const idx = args.indexOf(`--${name}`);
    return idx >= 0 && idx + 1 < args.length ? args[idx + 1] : undefined;
}

switch (command) {
    case "extract":
        const formId = getArg("form");
        if (!formId) { console.error("Usage: extract --form <formId> [--output <file.json>]"); process.exit(1); }
        cmdExtract(formId, getArg("output"));
        break;

    case "list-forms":
        const entity = getArg("entity");
        if (!entity) { console.error("Usage: list-forms --entity <entityName>"); process.exit(1); }
        cmdListForms(entity);
        break;

    case "apply": {
        const applyFormId = getArg("form");
        const applyConfig = getArg("config");
        const applyField = getArg("field");
        if (!applyFormId || !applyConfig) {
            console.error("Usage: apply --form <formId> --config <config.json> [--field <fieldName>]");
            process.exit(1);
        }
        cmdApply(applyFormId, applyConfig, applyField);
        break;
    }

    case "to-harness": {
        const configFile = getArg("config");
        if (!configFile) { console.error("Usage: to-harness --config <file.json> [--output <scenario.json>]"); process.exit(1); }
        const config: InfoCardConfig = JSON.parse(fs.readFileSync(configFile, "utf-8"));
        // Convert to harness test-scenario format using the phone factor (or first available)
        const params = config.formFactors.phone ?? config.formFactors.tablet ?? config.formFactors.desktop;
        if (!params) { console.error("No form factor found in config"); process.exit(1); }
        const propertyValues: Record<string, string> = {};
        for (const [name, binding] of Object.entries(params)) {
            if (binding.static) {
                propertyValues[name] = binding.value;
            } else if (binding.value.startsWith("@")) {
                propertyValues[name] = binding.value;
            } else {
                // Bound column — use $ prefix for harness convention
                propertyValues[name] = `$${binding.value}`;
            }
        }
        const scenario = {
            name: `Imported: ${config.extractedFrom.fieldName} (${config.extractedFrom.formId.substring(0, 8)})`,
            description: `Extracted from form ${config.extractedFrom.formId} on ${config.extractedFrom.extractedAt}`,
            savedAt: new Date().toISOString(),
            propertyValues,
            pageEntityTypeName: config.extractedFrom.entityName || "bookableresourcebooking",
            networkMode: "online",
            devicePreset: "iphone-14",
            isControlDisabled: false,
        };
        const out = getArg("output");
        const json = JSON.stringify(scenario, null, 2);
        if (out) { fs.writeFileSync(out, json, "utf-8"); console.log(`Harness scenario saved to ${out}`); }
        else { console.log(json); }
        break;
    }

    case "from-harness": {
        const scenarioFile = getArg("scenario");
        if (!scenarioFile) { console.error("Usage: from-harness --scenario <file.json> [--output <config.json>]"); process.exit(1); }
        const scenario = JSON.parse(fs.readFileSync(scenarioFile, "utf-8"));
        const pvs = scenario.propertyValues as Record<string, string>;
        const params: Record<string, ParameterBinding> = {};
        // Known config properties with their types
        const configTypes: Record<string, string> = {
            layout: "Enum", hideEmptyFields: "TwoOptions", showCardBorder: "TwoOptions",
            showVersionInfo: "TwoOptions", startExpanded: "TwoOptions", showTitle: "TwoOptions",
        };
        for (const [name, value] of Object.entries(pvs)) {
            if (!value && value !== "false") continue;
            if (configTypes[name]) {
                params[name] = { type: configTypes[name], value, static: true };
            } else if (value.startsWith("@")) {
                params[name] = { type: "SingleLine.Text", value, static: true };
            } else if (value.startsWith("$")) {
                // Determine type from property name heuristics
                let type = "Currency"; // default for lookups
                if (name.includes("latitude") || name.includes("longitude")) type = "FP";
                else if (name.includes("grid") && name.match(/\d/)) type = "DateAndTime.DateAndTime";
                else if (name.includes("tag")) type = "Lookup.Simple";
                params[name] = { type, value: value.substring(1) };
            }
        }
        const config: InfoCardConfig = {
            controlName: "smp_Sample.InfoCard",
            extractedFrom: {
                formId: "", formName: scenario.name, entityName: scenario.pageEntityTypeName ?? "",
                fieldName: pvs.titleField?.startsWith("$") ? pvs.titleField.substring(1) : "",
                extractedAt: new Date().toISOString(),
            },
            formFactors: { phone: params, tablet: params, desktop: params },
        };
        const out = getArg("output");
        const json = JSON.stringify(config, null, 2);
        if (out) { fs.writeFileSync(out, json, "utf-8"); console.log(`Config saved to ${out}`); }
        else { console.log(json); }
        break;
    }

    default:
        console.log(`
InfoCard Configuration Tool

Commands:
  extract      Extract InfoCard config from a Dataverse form → JSON
  to-harness   Convert config JSON → harness test-scenario format
  from-harness Convert harness scenario → config JSON (for apply)
  apply        Apply config JSON to a Dataverse form (coming soon)
  list-forms   List main forms for an entity

Options:
  --form <formId>        Form GUID (for extract)
  --entity <name>        Entity logical name (for list-forms)
  --config <file.json>   Config file (for to-harness, apply)
  --scenario <file.json> Scenario file (for from-harness)
  --output <file.json>   Output file

Workflows:
  # Extract live form → test in harness
  npx ts-node tools/infocard-config.ts extract --form 01633f27-... --output config.json
  npx ts-node tools/infocard-config.ts to-harness --config config.json --output scenario.json

  # Harness scenario → apply to another form
  npx ts-node tools/infocard-config.ts from-harness --scenario scenario.json --output config.json
  npx ts-node tools/infocard-config.ts apply --form <target-form-id> --config config.json
        `);
}
