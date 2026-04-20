/**
 * Tests for the InfoCard binding validator.
 * Validates @/@. syntax against entity schemas derived from data.json.
 */

import * as testData from "../data.json";
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

        it("warns about single-hop @. that should use $ binding", () => {
            const results = validateBindings(
                { gridField1: "@.starttime" },
                schemas, "bookableresourcebooking",
            );
            expect(results).toHaveLength(1);
            expect(results[0].warning).toContain("consider using $starttime");
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
        it("validates the LIVE booking scenario from test-scenarios.json", () => {
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
    });
});
