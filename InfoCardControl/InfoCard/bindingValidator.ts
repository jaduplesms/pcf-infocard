/**
 * InfoCard Binding Validator
 *
 * Validates @/@. syntax bindings against entity schemas from data.json.
 * Used by the harness at design time and available for unit tests.
 *
 * Validation rules:
 * 1. @ fields must exist on the title entity (resolved from the title lookup)
 * 2. @. fields must exist on the form entity (current record)
 * 3. Dotted paths (nav.field) must have a valid lookup nav property and a valid field on the target entity
 * 4. Syntax: @ must be followed by a field name, @. must be followed by nav.field
 */

/**
 * Common Dataverse columns whose `of-type` is NOT supported by the PCF manifest
 * (https://learn.microsoft.com/.../property#value-elements-that-arent-supported).
 * For these columns @./@ syntax is the *correct* path — the form designer's
 * column picker won't expose them for direct binding. Used to suppress the
 * "consider $-binding" warning so we don't mislead makers.
 */
export function isUnbindableColumn(col: string): boolean {
    if (!col) return false;
    const c = col.toLowerCase();
    // Exact name matches (Customer/Owner/Status/Regarding lookups, language/timezone)
    if (["customerid", "ownerid", "owninguser", "owningteam", "owningbusinessunit",
         "statecode", "statuscode", "regardingobjectid", "transactioncurrencyid"].includes(c)) {
        return true;
    }
    // Whole.Duration: any column ending in "duration" or "durationminutes"
    if (/(^|_)duration(minutes)?$/.test(c)) return true;
    // Whole.Language / Whole.TimeZone heuristics
    if (/(^|_)languagecode$/.test(c)) return true;
    if (/(^|_)timezonecode$/.test(c)) return true;
    return false;
}

export interface BindingValidation {
    slotKey: string;
    expression: string;
    type: "title-related" | "current-related" | "bound";
    valid: boolean;
    error?: string;
    warning?: string;
}

export interface EntitySchema {
    /** Column logical names that exist as direct primitives or lookups */
    columns: string[];
    /** Lookup columns: logicalName → target entity type */
    lookups: Record<string, string>;
}

/**
 * Build entity schemas from data.json records.
 * Infers columns from the first record of each entity type.
 */
export function buildSchemaFromData(
    data: Record<string, Array<Record<string, unknown>>>,
): Record<string, EntitySchema> {
    const schemas: Record<string, EntitySchema> = {};

    for (const [entityType, records] of Object.entries(data)) {
        if (!Array.isArray(records) || records.length === 0) continue;
        const record = records[0];
        const columns: string[] = [];
        const lookups: Record<string, string> = {};

        for (const key of Object.keys(record)) {
            if (key.includes("@")) continue; // skip OData annotations

            if (key.startsWith("_") && key.endsWith("_value")) {
                // Lookup column: _field_value → field
                const logicalName = key.slice(1, -6); // remove _ prefix and _value suffix
                columns.push(logicalName);
                // Get target entity from annotation
                const etnKey = `${key}@Microsoft.Dynamics.CRM.lookuplogicalname`;
                if (record[etnKey]) {
                    lookups[logicalName] = String(record[etnKey]);
                }
            } else {
                columns.push(key);
            }
        }

        schemas[entityType] = { columns, lookups };
    }

    return schemas;
}

/**
 * Validate all @ bindings in a test scenario's property values.
 */
export function validateBindings(
    propertyValues: Record<string, string>,
    schemas: Record<string, EntitySchema>,
    formEntity: string,
    titleEntity?: string,
): BindingValidation[] {
    const results: BindingValidation[] = [];

    for (const [slotKey, value] of Object.entries(propertyValues)) {
        if (!value || typeof value !== "string") continue;

        // Bound column ($ prefix in harness convention)
        if (value.startsWith("$")) {
            const col = value.substring(1);
            const schema = schemas[formEntity];
            const valid = schema ? schema.columns.includes(col) : false;
            results.push({
                slotKey,
                expression: value,
                type: "bound",
                valid,
                error: !valid && schema ? `Column '${col}' not found on ${formEntity}` : undefined,
            });
            continue;
        }

        // @. prefix: current record navigation
        if (value.startsWith("@.")) {
            const path = value.substring(2).trim();
            if (!path) {
                results.push({ slotKey, expression: value, type: "current-related", valid: false, error: "Empty path after @." });
                continue;
            }

            const dotIdx = path.indexOf(".");
            if (dotIdx <= 0) {
                // Single-hop on current record — usually $-binding is faster, but
                // skip the warning entirely for known unbindable column types
                // (Whole.Duration / Status / Customer / Owner / Regarding / etc.)
                // where @. is the *required* path.
                const schema = schemas[formEntity];
                const exists = schema ? schema.columns.includes(path) : false;
                const unbindable = isUnbindableColumn(path);
                results.push({
                    slotKey, expression: value, type: "current-related",
                    valid: exists,
                    warning: (exists && !unbindable)
                        ? `'${path}' is a direct column — consider $${path} binding (faster, design-time validation). Stay with @.${path} if the column type isn't bindable.`
                        : undefined,
                    error: !exists ? `Column '${path}' not found on ${formEntity}` : undefined,
                });
                continue;
            }

            // Dotted path: nav.field
            const navProp = path.substring(0, dotIdx);
            const field = path.substring(dotIdx + 1);
            const schema = schemas[formEntity];

            if (!schema) {
                results.push({ slotKey, expression: value, type: "current-related", valid: false, error: `Entity '${formEntity}' not in schema` });
                continue;
            }

            const targetEntity = schema.lookups[navProp];
            if (!targetEntity) {
                const isColumn = schema.columns.includes(navProp);
                results.push({
                    slotKey, expression: value, type: "current-related", valid: false,
                    error: isColumn
                        ? `'${navProp}' exists on ${formEntity} but is not a lookup — cannot navigate to .${field}`
                        : `Lookup '${navProp}' not found on ${formEntity}`,
                });
                continue;
            }

            const targetSchema = schemas[targetEntity];
            const fieldExists = targetSchema ? targetSchema.columns.includes(field) : true; // unknown entity = assume valid
            results.push({
                slotKey, expression: value, type: "current-related",
                valid: fieldExists,
                error: !fieldExists ? `Column '${field}' not found on ${targetEntity} (via ${navProp})` : undefined,
            });
            continue;
        }

        // @ prefix: title entity navigation
        if (value.startsWith("@")) {
            const path = value.substring(1).trim();
            if (!path) {
                results.push({ slotKey, expression: value, type: "title-related", valid: false, error: "Empty path after @" });
                continue;
            }

            if (!titleEntity) {
                results.push({ slotKey, expression: value, type: "title-related", valid: false, error: "Title entity unknown — cannot validate @ reference" });
                continue;
            }

            const dotIdx = path.indexOf(".");
            if (dotIdx <= 0) {
                // Direct field on title entity
                const schema = schemas[titleEntity];
                const exists = schema ? schema.columns.includes(path) : true;
                results.push({
                    slotKey, expression: value, type: "title-related",
                    valid: exists,
                    error: !exists ? `Column '${path}' not found on ${titleEntity}` : undefined,
                });
                continue;
            }

            // Dotted path: nav.field on title entity
            const navProp = path.substring(0, dotIdx);
            const field = path.substring(dotIdx + 1);
            const schema = schemas[titleEntity];

            if (!schema) {
                results.push({ slotKey, expression: value, type: "title-related", valid: false, error: `Entity '${titleEntity}' not in schema` });
                continue;
            }

            const targetEntity = schema.lookups[navProp];
            if (!targetEntity) {
                results.push({
                    slotKey, expression: value, type: "title-related", valid: false,
                    error: `Lookup '${navProp}' not found on ${titleEntity}`,
                });
                continue;
            }

            const targetSchema = schemas[targetEntity];
            const fieldExists = targetSchema ? targetSchema.columns.includes(field) : true;
            results.push({
                slotKey, expression: value, type: "title-related",
                valid: fieldExists,
                error: !fieldExists ? `Column '${field}' not found on ${targetEntity} (via ${navProp})` : undefined,
            });
            continue;
        }
    }

    return results;
}

/**
 * Resolve the title entity from the title field binding.
 * In harness: $msdyn_workorder → look up the lookup target for msdyn_workorder on the form entity.
 */
export function resolveTitleEntity(
    titleBinding: string,
    formEntity: string,
    schemas: Record<string, EntitySchema>,
): string | undefined {
    if (!titleBinding.startsWith("$")) return undefined;
    const col = titleBinding.substring(1);
    const schema = schemas[formEntity];
    return schema?.lookups[col];
}
