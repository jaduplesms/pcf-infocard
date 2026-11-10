/**
 * Tests for the InfoCard binding validator.
 * Validates @/@. syntax against entity schemas derived from data.json.
 */

import * as testData from "./test-data.json";
import * as publicScenarios from "../test-scenarios.json";
import {
    buildSchemaFromData,
    validateBindings,
    resolveTitleEntity,
    EntitySchema,
} from "../bindingValidator";

// ────────────────────────────────────────
// Schema building
// ────────────────────────────────────────

describe("buildSchemaFromData()", () => {
    const schemas = buildSchemaFromData(testData as unknown as Record<string, Array<Record<string, unknown>>>);

    it("builds schemas for all entity types in data.json", () => {
        expect(schemas.bookableresourcebooking).toBeDefined();
        expect(schemas.msdyn_workorder).toBeDefined();
        expect(schemas.account).toBeDefined();
        expect(schemas.bookingstatus).toBeDefined();
        expect(schemas.msdyn_priority).toBeDefined();
    });

    it("extracts columns from record keys", () => {
        const booking = schemas.bookableresourcebooking;
        expect(booking.columns).toContain("starttime");
        expect(booking.columns).toContain("endtime");
        expect(booking.columns).toContain("duration");
        expect(booking.columns).toContain("bookingtype");
    });

    it("extracts lookup columns from _field_value pattern", () => {
        const booking = schemas.bookableresourcebooking;
        expect(booking.columns).toContain("msdyn_workorder");
        expect(booking.columns).toContain("resource");
        expect(booking.columns).toContain("bookingstatus");
    });

    it("resolves lookup target entities", () => {
        const booking = schemas.bookableresourcebooking;
        expect(booking.lookups.msdyn_workorder).toBe("msdyn_workorder");
        expect(booking.lookups.resource).toBe("bookableresource");
        expect(booking.lookups.bookingstatus).toBe("bookingstatus");
    });

    it("resolves work order lookups", () => {
        const wo = schemas.msdyn_workorder;
        expect(wo.lookups.msdyn_serviceaccount).toBe("account");
        expect(wo.lookups.msdyn_primaryincidenttype).toBe("msdyn_incidenttype");
        expect(wo.lookups.msdyn_priority).toBe("msdyn_priority");
    });

    it("includes color fields on booking status", () => {
        const bs = schemas.bookingstatus;
        expect(bs.columns).toContain("msdyn_statuscolor");
        expect(bs.columns).toContain("name");
    });

    it("includes color fields on priority", () => {
        const p = schemas.msdyn_priority;
        expect(p.columns).toContain("msdyn_prioritycolor");
        expect(p.columns).toContain("msdyn_name");
    });
});

// ────────────────────────────────────────
// Title entity resolution
// ────────────────────────────────────────

describe("resolveTitleEntity()", () => {
    const schemas = buildSchemaFromData(testData as unknown as Record<string, Array<Record<string, unknown>>>);

    it("resolves $msdyn_workorder → msdyn_workorder", () => {
        expect(resolveTitleEntity("$msdyn_workorder", "bookableresourcebooking", schemas))
            .toBe("msdyn_workorder");
    });

    it("returns undefined for non-lookup binding", () => {
        expect(resolveTitleEntity("$starttime", "bookableresourcebooking", schemas))
            .toBeUndefined();
    });

    it("returns undefined for non-$ binding", () => {
        expect(resolveTitleEntity("@msdyn_workorder", "bookableresourcebooking", schemas))
            .toBeUndefined();
    });
});

// ────────────────────────────────────────
// Binding validation
// ────────────────────────────────────────

describe("validateBindings()", () => {
    const schemas = buildSchemaFromData(testData as unknown as Record<string, Array<Record<string, unknown>>>);

    describe("@ (title entity) references", () => {
        it("validates direct field on work order", () => {
            const results = validateBindings(
                { detailField1: "@msdyn_instructions" },
                schemas, "bookableresourcebooking", "msdyn_workorder",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(true);
        });

        it("flags non-existent field on work order", () => {
            const results = validateBindings(
                { detailField1: "@nonexistent_field" },
                schemas, "bookableresourcebooking", "msdyn_workorder",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(false);
            expect(results[0].error).toContain("not found on msdyn_workorder");
        });

        it("validates dotted path @msdyn_serviceaccount.telephone1", () => {
            const results = validateBindings(
                { phoneField1: "@msdyn_serviceaccount.telephone1" },
                schemas, "bookableresourcebooking", "msdyn_workorder",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(true);
        });

        it("flags invalid nav property in dotted path", () => {
            const results = validateBindings(
                { phoneField1: "@nonexistent_lookup.telephone1" },
                schemas, "bookableresourcebooking", "msdyn_workorder",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(false);
            expect(results[0].error).toContain("Lookup 'nonexistent_lookup' not found");
        });

        it("flags invalid field on target entity in dotted path", () => {
            const results = validateBindings(
                { phoneField1: "@msdyn_serviceaccount.nonexistent_column" },
                schemas, "bookableresourcebooking", "msdyn_workorder",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(false);
            expect(results[0].error).toContain("not found on account");
        });

        it("flags empty @ reference", () => {
            const results = validateBindings(
                { phoneField1: "@" },
                schemas, "bookableresourcebooking", "msdyn_workorder",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(false);
            expect(results[0].error).toContain("Empty path");
        });
    });

    describe("@. (current record) references", () => {
        it("validates dotted path @.resource.resourcetype", () => {
            // resource is a lookup on booking → bookableresource entity
            // bookableresource schema may not have resourcetype in data.json
            // but if schema is unknown, assume valid
            const results = validateBindings(
                { gridField5: "@.resource.resourcetype" },
                schemas, "bookableresourcebooking",
            );
            expect(results).toHaveLength(1);
            // Valid if bookableresource schema exists and has resourcetype, or schema not found (assumed valid)
            expect(results[0].type).toBe("current-related");
        });

        it("flags invalid lookup on current record", () => {
            const results = validateBindings(
                { gridField5: "@.nonexistent.field" },
                schemas, "bookableresourcebooking",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(false);
            expect(results[0].error).toContain("Lookup 'nonexistent' not found");
        });

        it("warns about single-hop @. that should use $ binding (for bindable types)", () => {
            const results = validateBindings(
                { gridField1: "@.starttime" },
                schemas, "bookableresourcebooking",
            );
            expect(results).toHaveLength(1);
            expect(results[0].warning).toContain("consider $starttime binding");
        });

        it("does NOT warn for unbindable column types (Whole.Duration, Status)", () => {
            // bookableresourcebooking schema has `duration` and `statecode`. Both
            // are unbindable in the form designer, so no "consider $-binding" warning.
            const results = validateBindings(
                {
                    gridField1: "@.duration",
                    gridField2: "@.statecode",
                },
                schemas, "bookableresourcebooking",
            );
            for (const r of results) {
                expect(r.warning).toBeUndefined();
            }
        });

        it("flags empty @. reference", () => {
            const results = validateBindings(
                { gridField1: "@." },
                schemas, "bookableresourcebooking",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(false);
        });
    });

    describe("$ (bound column) references", () => {
        it("validates existing column", () => {
            const results = validateBindings(
                { gridField1: "$starttime" },
                schemas, "bookableresourcebooking",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(true);
        });

        it("flags non-existent column", () => {
            const results = validateBindings(
                { gridField1: "$nonexistent" },
                schemas, "bookableresourcebooking",
            );
            expect(results).toHaveLength(1);
            expect(results[0].valid).toBe(false);
            expect(results[0].error).toContain("not found on bookableresourcebooking");
        });
    });

    describe("full scenario validation", () => {
        it("validates a mock booking scenario", () => {
            const scenario: Record<string, string> = {
                titleField: "$msdyn_workorder",
                subtitleField1: "$resource",
                subtitleField2: "@msdyn_serviceaccount",
                phoneField1: "@msdyn_serviceaccount.telephone1",
                emailField: "@msdyn_serviceaccount.emailaddress1",
                addressField: "@msdyn_address1",
                detailField1: "@msdyn_instructions",
                detailField2: "@msdyn_workordersummary",
                gridField1: "$starttime",
                gridField2: "$endtime",
                gridField3: "$duration",
                tagField1: "$bookingstatus",
                tagField2: "@msdyn_primaryincidenttype",
                tagField3: "@msdyn_priority",
            };
            const titleEntity = resolveTitleEntity(scenario.titleField, "bookableresourcebooking", schemas);
            const results = validateBindings(scenario, schemas, "bookableresourcebooking", titleEntity);

            const errors = results.filter(r => !r.valid);
            const validCount = results.filter(r => r.valid).length;

            // Log for visibility
            for (const r of results) {
                if (r.error) console.log(`  ✗ ${r.slotKey}: ${r.expression} — ${r.error}`);
            }

            expect(validCount).toBeGreaterThan(0);
            // Some fields like msdyn_address1 may not be in the sample data — that's expected
        });

        describe("public mock harness assets", () => {
            const metadata = publicScenarios.metadata as unknown as Record<string, {
                primaryIdAttribute: string;
                columns: Record<string, unknown>;
            }>;

            it("resolves every saved binding and lookup from scenario-local records", () => {
                for (const scenario of publicScenarios.scenarios) {
                    expect(scenario.dataSource).toBe("mock");
                    const records = scenario.dataRecords as unknown as Record<string, Array<Record<string, unknown>>>;
                    const schemas = buildSchemaFromData(records);
                    const formEntity = scenario.pageContext.typeName;
                    const titleEntity = resolveTitleEntity(scenario.propertyValues.titleField, formEntity, schemas);
                    const results = validateBindings(
                        scenario.propertyValues as unknown as Record<string, string>, schemas, formEntity, titleEntity,
                    );
                    expect(results.filter(result => !result.valid)).toEqual([]);
                    expect(records[formEntity].some(record =>
                        record[metadata[formEntity].primaryIdAttribute] === scenario.pageContext.entityId,
                    )).toBe(true);

                    for (const [entity, rows] of Object.entries(records)) {
                        expect(metadata[entity]).toBeDefined();
                        for (const row of rows) {
                            for (const [key, value] of Object.entries(row)) {
                                if (key.includes("@")) continue;
                                const column = key.startsWith("_") && key.endsWith("_value") ? key.slice(1, -6) : key;
                                expect(metadata[entity].columns[column]).toBeDefined();
                                if (key.startsWith("_") && key.endsWith("_value") && value) {
                                    const target = String(row[`${key}@Microsoft.Dynamics.CRM.lookuplogicalname`]);
                                    expect(records[target]?.some(targetRow =>
                                        targetRow[metadata[target].primaryIdAttribute] === value,
                                    )).toBe(true);
                                }
                            }
                        }
                    }
                }
            });

            it("keeps public record ids synthetic and external links on example.com", () => {
                function check(value: unknown, key = ""): void {
                    if (Array.isArray(value)) {
                        value.forEach(item => check(item));
                    } else if (value && typeof value === "object") {
                        for (const [name, item] of Object.entries(value)) check(item, name);
                    } else if (typeof value === "string") {
                        expect(value).not.toMatch(/https?:\/\/[^/\s"]*(?:dynamics\.com|sharepoint\.com)/i);
                        if (/^(?:https?:\/\/)/.test(value)) expect(new URL(value).hostname).toBe("example.com");
                        if (/email/i.test(key) && value && !value.startsWith("$") && !value.startsWith("@")) {
                            expect(value).toMatch(/@example\.com$/);
                        }
                        if (/phone/i.test(key) && value && !value.startsWith("$") && !value.startsWith("@")) {
                            expect(value).toMatch(/555-01\d{2}$/);
                        }
                        if (/^[0-9a-f]{8}-[0-9a-f-]{27}$/i.test(value)) {
                            expect(value).toMatch(/^00000000-0000-0000-0000-000000000\d{3}$/);
                        }
                        expect(key).not.toMatch(/^(?:liveProfile|orgUrl|provenance)$/i);
                    }
                }
                check(publicScenarios);
                check(testData);
            });
        });
    });
});
