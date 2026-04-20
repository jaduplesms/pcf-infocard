/**
 * Unit tests for the InfoCard PCF lifecycle class (index.ts).
 *
 * Covers: init, updateView, collectData, readSlot, readSlotGroup,
 *         resolveLayout, resolveTheme, detectRelatedFields,
 *         isDurationField, formatDuration, getOutputs, destroy.
 */

import { InfoCard } from "../index";
import { IInputs } from "../generated/ManifestTypes";
import { defaultTheme } from "../InfoCard";
import * as testData from "../data.json";

// ────────────────────────────────────────
// Mock context helper
// ────────────────────────────────────────

interface SlotParamOverride {
  raw?: unknown;
  formatted?: string;
  type?: string;
  attributes?: Record<string, unknown>;
}

interface ContextOverrides {
  layout?: string | null;
  hideEmptyFields?: boolean;
  showCardBorder?: boolean;
  showVersionInfo?: boolean;
  relatedConfig?: string | null;
  fluentDesignLanguage?: Record<string, unknown>;
  contextInfo?: { entityTypeName: string; entityId: string };
  slots?: Record<string, SlotParamOverride>;
  /** Custom WebAPI mock — overrides the default empty retrieveRecord */
  webAPIMock?: jest.Mock;
  /** Custom metadata mock — overrides the default empty getEntityMetadata */
  metadataMock?: jest.Mock;
}

/**
 * Simulates the PCF runtime param object.
 *
 * IMPORTANT: In the real Dynamics 365 runtime, only `usage="bound"` properties
 * (titleField) get attributes with DisplayName/LogicalName. All `usage="input"`
 * properties (grids, tags, details, subtitles, etc.) get attributes={}.
 * The `bound` parameter controls this behavior.
 */
function makeSlotParam(override?: SlotParamOverride, bound = false) {
  if (!override) {
    return { raw: null, type: "Unknown", formatted: undefined, attributes: undefined };
  }
  if (override.type === "Unknown") {
    return { raw: override.raw ?? null, type: "Unknown", formatted: undefined, attributes: undefined };
  }
  // Bound properties get full attributes from the platform.
  // Input properties get empty attributes (like the real runtime).
  const defaultAttrs = bound
    ? { LogicalName: "field", DisplayName: "Field" }
    : {};
  return {
    raw: override.raw ?? null,
    type: override.type ?? "SingleLine.Text",
    formatted: override.formatted,
    attributes: override.attributes ?? defaultAttrs,
  };
}

function createMockContext(overrides: ContextOverrides = {}): ComponentFramework.Context<IInputs> {
  const slots = overrides.slots ?? {};

  const allSlotKeys = [
    "titleField",
    "subtitleField1", "subtitleField2", "subtitleField3",
    "phoneField1", "phoneField2",
    "emailField", "webField",
    "addressField", "latitudeField", "longitudeField",
    "detailField1", "detailField2", "detailField3", "detailField4", "detailField5",
    "gridField1", "gridField2", "gridField3", "gridField4", "gridField5",
    "gridField6", "gridField7", "gridField8", "gridField9", "gridField10",
    "tagField1", "tagField2", "tagField3", "tagField4", "tagField5",
  ];

  const parameters: Record<string, unknown> = {
    layout: { raw: overrides.layout ?? "contact" },
    hideEmptyFields: { raw: overrides.hideEmptyFields ?? true },
    showCardBorder: { raw: overrides.showCardBorder ?? true },
    showVersionInfo: { raw: overrides.showVersionInfo ?? false },
    showTitle: { raw: true },
    startExpanded: { raw: true },
    relatedConfig: { raw: overrides.relatedConfig ?? null },
  };

  for (const key of allSlotKeys) {
    // Only titleField is usage="bound" — gets attributes from the platform
    const isBound = key === "titleField";
    parameters[key] = makeSlotParam(slots[key], isBound);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const modeObj: Record<string, any> = {
    isControlDisabled: false,
    allocatedWidth: 350,
    allocatedHeight: 500,
    isVisible: true,
    label: "InfoCard",
  };
  if (overrides.contextInfo) {
    modeObj.contextInfo = overrides.contextInfo;
  }

  const ctx: Record<string, unknown> = {
    parameters,
    mode: modeObj as ComponentFramework.Mode,
    client: {} as ComponentFramework.Client,
    device: {} as ComponentFramework.Device,
    factory: {} as ComponentFramework.Factory,
    formatting: {} as ComponentFramework.Formatting,
    navigation: {
      openForm: jest.fn().mockResolvedValue(undefined),
    } as unknown as ComponentFramework.Navigation,
    resources: {} as ComponentFramework.Resources,
    userSettings: {} as ComponentFramework.UserSettings,
    utils: {
      getEntityMetadata: overrides.metadataMock ?? jest.fn().mockResolvedValue({
        Attributes: { getAll: () => [] },
      }),
    } as unknown as ComponentFramework.Utility,
    webAPI: {
      retrieveRecord: overrides.webAPIMock ?? jest.fn().mockResolvedValue({}),
    } as unknown as ComponentFramework.WebApi,
    accessibility: undefined,
    events: undefined,
    theming: undefined,
    page: undefined,
    orgSettings: undefined,
  };

  if (overrides.fluentDesignLanguage) {
    ctx.fluentDesignLanguage = overrides.fluentDesignLanguage;
  }

  return ctx as unknown as ComponentFramework.Context<IInputs>;
}

/** Create a realistic WebAPI record with getFormattedValue method */
function makeWebAPIRecord(data: Record<string, unknown>): Record<string, unknown> {
  return {
    ...data,
    getFormattedValue(col: string) {
      return (data as Record<string, unknown>)[`${col}@OData.Community.Display.V1.FormattedValue`] ?? null;
    },
  };
}

/** Create a realistic entity metadata response */
function makeMetadataResponse(attributes: Array<{
  LogicalName: string;
  DisplayName: string;
  options?: Array<{ Value: number; Label: string; Color?: string }>;
}>) {
  return {
    Attributes: {
      getAll: () => attributes.map(a => ({
        LogicalName: a.LogicalName,
        DisplayName: { UserLocalizedLabel: { Label: a.DisplayName } },
        attributeDescriptor: a.options ? {
          OptionSet: {
            Options: a.options.map(o => ({
              Value: o.Value,
              Label: { UserLocalizedLabel: { Label: o.Label } },
              Color: o.Color,
            })),
          },
        } : undefined,
      })),
    },
  };
}

/** The booking record from data.json, wrapped as a WebAPI entity */
const BOOKING_RECORD = makeWebAPIRecord(
  testData.bookableresourcebooking[0] as unknown as Record<string, unknown>
);

// Utility to access private methods via any-cast
function asAny(obj: unknown): Record<string, (...args: unknown[]) => unknown> {
  return obj as Record<string, (...args: unknown[]) => unknown>;
}

// ────────────────────────────────────────
// Tests
// ────────────────────────────────────────

describe("InfoCard PCF Lifecycle", () => {
  let control: InfoCard;
  let notifyOutputChanged: jest.Mock;

  beforeEach(() => {
    control = new InfoCard();
    notifyOutputChanged = jest.fn();
  });

  // ── init ──────────────────────────────

  describe("init()", () => {
    it("stores context", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      // Context stored — updateView should work without error
      const element = control.updateView(ctx);
      expect(element).toBeDefined();
    });
  });

  // ── updateView ────────────────────────

  describe("updateView()", () => {
    it("returns a React element", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Test Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element).toBeDefined();
      expect(element.type).toBeDefined();
      expect(element.props).toBeDefined();
    });

    it("passes hideEmpty, showBorder, showVersion from context params", () => {
      const ctx = createMockContext({
        hideEmptyFields: false,
        showCardBorder: false,
        showVersionInfo: true,
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.hideEmpty).toBe(false);
      expect(element.props.showBorder).toBe(false);
      expect(element.props.showVersion).toBe(true);
    });
  });

  // ── collectData (via updateView props) ─

  describe("collectData()", () => {
    it("reads title field", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Work Order 123",
            attributes: { LogicalName: "msdyn_name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.title).not.toBeNull();
      expect(element.props.data.title.value).toBe("Work Order 123");
      expect(element.props.data.title.label).toBe("Name");
      expect(element.props.data.title.isEmpty).toBe(false);
    });

    it("reads subtitle fields", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          subtitleField1: {
            raw: "Contoso",
            attributes: { LogicalName: "account", DisplayName: "Account" },
          },
          subtitleField2: {
            raw: "High",
            attributes: { LogicalName: "priority", DisplayName: "Priority" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.subtitles).toHaveLength(2);
      expect(element.props.data.subtitles[0].value).toBe("Contoso");
      expect(element.props.data.subtitles[1].value).toBe("High");
    });

    it("reads phone fields", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          phoneField1: {
            raw: "+1 555 0123",
            attributes: { LogicalName: "phone1", DisplayName: "Phone" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.phones).toHaveLength(1);
      expect(element.props.data.phones[0].value).toBe("+1 555 0123");
    });

    it("reads email field", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          emailField: {
            raw: "test@example.com",
            attributes: { LogicalName: "email", DisplayName: "Email" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.email).not.toBeNull();
      expect(element.props.data.email.value).toBe("test@example.com");
    });

    it("reads detail fields", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          detailField1: {
            raw: "123 Main St",
            attributes: { LogicalName: "address1", DisplayName: "Address" },
          },
          detailField2: {
            raw: "Some notes here",
            attributes: { LogicalName: "notes", DisplayName: "Notes" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.details).toHaveLength(2);
      expect(element.props.data.details[0].value).toBe("123 Main St");
    });

    it("reads grid fields", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          gridField1: {
            raw: "Value A",
            attributes: { LogicalName: "fieldA", DisplayName: "Field A" },
          },
          gridField2: {
            raw: "Value B",
            attributes: { LogicalName: "fieldB", DisplayName: "Field B" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.gridFields).toHaveLength(2);
      expect(element.props.data.gridFields[0].label).toBe("Field A");
    });

    it("reads tag fields", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          tagField1: {
            raw: "Urgent",
            attributes: { LogicalName: "tag1", DisplayName: "Tag 1" },
          },
          tagField3: {
            raw: "Open",
            attributes: { LogicalName: "tag3", DisplayName: "Status" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.tags).toHaveLength(2);
      expect(element.props.data.tags[0].value).toBe("Urgent");
      expect(element.props.data.tags[1].value).toBe("Open");
    });

    it("reads latitude and longitude as numbers", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          latitudeField: {
            raw: "47.6062",
            attributes: { LogicalName: "lat", DisplayName: "Latitude" },
          },
          longitudeField: {
            raw: "-122.3321",
            attributes: { LogicalName: "lng", DisplayName: "Longitude" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.latitude).toBe(47.6062);
      expect(element.props.data.longitude).toBe(-122.3321);
    });

    it("sets lat/lng to null when not provided", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.latitude).toBeNull();
      expect(element.props.data.longitude).toBeNull();
    });
  });

  // ── readSlot ──────────────────────────

  describe("readSlot()", () => {
    function getReadSlot(ctrl: InfoCard) {
      return asAny(ctrl).readSlot.bind(ctrl) as (
        ctx: ComponentFramework.Context<IInputs>,
        key: string,
      ) => { label: string; value: string; rawValue: unknown; isEmpty: boolean; lookupEntityType?: string; lookupId?: string } | null;
    }

    it("returns null when param is missing", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      // "nonExistent" is not in parameters
      const result = readSlot(ctx, "nonExistent");
      expect(result).toBeNull();
    });

    it("returns null when param type is Unknown", () => {
      const ctx = createMockContext({
        slots: {
          titleField: { type: "Unknown" },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "titleField");
      expect(result).toBeNull();
    });

    it("returns null when attributes have no metadata and no type", () => {
      // Empty attributes + no type = improperly configured field
      const ctx = createMockContext();
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const params = ctx.parameters as Record<string, any>;
      params.titleField = { raw: null, type: undefined, formatted: undefined, attributes: {} };
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "titleField");
      expect(result).toBeNull();
    });

    it("returns field when attributes are empty but type is present (static input)", () => {
      // Static inputs (static="true" in form XML) have attributes={} but valid type
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "@msdyn_serviceaccount",
            attributes: {},
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "titleField");
      expect(result).not.toBeNull();
      expect(result!.isEmpty).toBe(true);
    });

    it("marks empty for null raw value", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: null,
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "titleField");

      expect(result).not.toBeNull();
      expect(result!.isEmpty).toBe(true);
      expect(result!.value).toBe("---");
    });

    it("uses formatted value when available", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: 3,
            formatted: "High",
            attributes: { LogicalName: "priority", DisplayName: "Priority" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "titleField");

      expect(result!.value).toBe("High");
      expect(result!.isEmpty).toBe(false);
    });

    it("handles lookup arrays (name present)", () => {
      const ctx = createMockContext({
        slots: {
          subtitleField1: {
            raw: [{ id: "abc-123", name: "Contoso Ltd", entityType: "account" }],
            attributes: { LogicalName: "customerid", DisplayName: "Customer" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "subtitleField1");

      expect(result!.value).toBe("Contoso Ltd");
      expect(result!.lookupEntityType).toBe("account");
      expect(result!.lookupId).toBe("abc-123");
      expect(result!.isEmpty).toBe(false);
    });

    it("handles lookup arrays (name missing, falls back to id)", () => {
      const ctx = createMockContext({
        slots: {
          subtitleField1: {
            raw: [{ id: "abc-123" }],
            attributes: { LogicalName: "lookupfield", DisplayName: "Lookup" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "subtitleField1");

      expect(result!.value).toBe("abc-123");
    });

    it("handles lookup arrays with etn instead of entityType", () => {
      const ctx = createMockContext({
        slots: {
          subtitleField1: {
            raw: [{ id: "def-456", name: "Resource", etn: "bookableresource" }],
            attributes: { LogicalName: "resource", DisplayName: "Resource" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "subtitleField1");

      expect(result!.lookupEntityType).toBe("bookableresource");
    });

    it("falls back to Targets attribute for lookup entity type", () => {
      const ctx = createMockContext({
        slots: {
          subtitleField1: {
            raw: [{ id: "abc-123", name: "WO-00001" }],
            attributes: {
              LogicalName: "msdyn_workorder",
              DisplayName: "Work Order",
              Targets: ["msdyn_workorder"],
            },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "subtitleField1");

      expect(result!.lookupEntityType).toBe("msdyn_workorder");
      expect(result!.lookupId).toBe("abc-123");
    });

    it("handles single lookup object (not array)", () => {
      const ctx = createMockContext({
        slots: {
          subtitleField1: {
            raw: { name: "Single Lookup" },
            attributes: { LogicalName: "ref", DisplayName: "Reference" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "subtitleField1");

      expect(result!.value).toBe("Single Lookup");
      expect(result!.isEmpty).toBe(false);
    });

    it("handles single lookup object without name", () => {
      const ctx = createMockContext({
        slots: {
          subtitleField1: {
            raw: { id: "no-name" },
            attributes: { LogicalName: "ref", DisplayName: "Reference" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "subtitleField1");

      expect(result!.value).toBe("---");
      expect(result!.isEmpty).toBe(true);
    });

    it("converts primitive raw to string", () => {
      const ctx = createMockContext({
        slots: {
          gridField1: {
            raw: 42,
            attributes: { LogicalName: "count", DisplayName: "Count" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "gridField1");

      expect(result!.value).toBe("42");
      expect(result!.isEmpty).toBe(false);
    });

    it("marks empty when displayValue is empty string", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "titleField");

      expect(result!.isEmpty).toBe(true);
    });

    it("formats duration fields", () => {
      const ctx = createMockContext({
        slots: {
          gridField1: {
            raw: 150,
            attributes: { LogicalName: "duration", DisplayName: "Duration", Format: "Duration" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "gridField1");

      expect(result!.value).toBe("2h 30m");
    });

    it("uses label from DisplayName, falls back to LogicalName", () => {
      const ctx1 = createMockContext({
        slots: {
          titleField: {
            raw: "val",
            attributes: { LogicalName: "logname", DisplayName: "Display Name" },
          },
        },
      });
      control.init(ctx1, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      expect(readSlot(ctx1, "titleField")!.label).toBe("Display Name");

      const ctx2 = createMockContext({
        slots: {
          titleField: {
            raw: "val",
            attributes: { LogicalName: "logname" },
          },
        },
      });
      // With no DisplayName, formatLogicalName is used: "logname" → "Logname"
      expect(readSlot(ctx2, "titleField")!.label).toBe("Logname");
    });
  });

  // ── readSlotGroup ─────────────────────

  describe("readSlotGroup()", () => {
    it("collects non-null slots and skips nulls", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          subtitleField1: {
            raw: "Sub1",
            attributes: { LogicalName: "s1", DisplayName: "S1" },
          },
          // subtitleField2 is Unknown (default) — should be skipped
          subtitleField3: {
            raw: "Sub3",
            attributes: { LogicalName: "s3", DisplayName: "S3" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.subtitles).toHaveLength(2);
      expect(element.props.data.subtitles[0].value).toBe("Sub1");
      expect(element.props.data.subtitles[1].value).toBe("Sub3");
    });

    it("returns empty array when no slots are configured", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.data.subtitles).toHaveLength(0);
      expect(element.props.data.phones).toHaveLength(0);
      expect(element.props.data.details).toHaveLength(0);
      expect(element.props.data.gridFields).toHaveLength(0);
      expect(element.props.data.tags).toHaveLength(0);
    });
  });

  // ── resolveLayout ─────────────────────

  describe("resolveLayout()", () => {
    it.each(["smart", "contact", "compact"] as const)(
      'returns "%s" when layout param is "%s"',
      (layout) => {
        const ctx = createMockContext({
          layout,
          slots: {
            titleField: {
              raw: "Title",
              attributes: { LogicalName: "name", DisplayName: "Name" },
            },
          },
        });
        control.init(ctx, notifyOutputChanged);
        const element = control.updateView(ctx);
        expect(element.props.layout).toBe(layout);
      },
    );

    it('defaults to "contact" for unrecognized value', () => {
      const ctx = createMockContext({
        layout: "invalid_layout",
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);
      expect(element.props.layout).toBe("contact");
    });

    it('defaults to "contact" for null layout', () => {
      const ctx = createMockContext({
        layout: null,
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);
      expect(element.props.layout).toBe("contact");
    });
  });

  // ── resolveTheme ──────────────────────

  describe("resolveTheme()", () => {
    it("returns defaultTheme when no fluentDesignLanguage", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.theme).toEqual(defaultTheme);
    });

    it("reads from fluentDesignLanguage tokens when available", () => {
      const ctx = createMockContext({
        fluentDesignLanguage: {
          tokenTheme: {
            colorNeutralBackground1: "#111",
            colorNeutralForeground1: "#222",
            colorNeutralForeground3: "#333",
            colorNeutralForeground4: "#444",
            colorNeutralStroke1: "#555",
            colorNeutralStroke2: "#666",
            colorBrandForeground1: "#777",
            colorBrandBackground2: "#888",
            borderRadiusMedium: "12px",
            shadow4: "0 0 4px red",
          },
        },
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.theme.cardBg).toBe("#111");
      expect(element.props.theme.textPrimary).toBe("#222");
      expect(element.props.theme.textSecondary).toBe("#333");
      expect(element.props.theme.textMuted).toBe("#444");
      expect(element.props.theme.border).toBe("#555");
      expect(element.props.theme.borderLight).toBe("#666");
      expect(element.props.theme.brand).toBe("#777");
      expect(element.props.theme.brandLight).toBe("#888");
      expect(element.props.theme.radius).toBe("12px");
      expect(element.props.theme.shadow).toBe("0 0 4px red");
      // fontFamily always comes from defaultTheme
      expect(element.props.theme.fontFamily).toBe(defaultTheme.fontFamily);
    });

    it("falls back to defaultTheme values for missing tokens", () => {
      const ctx = createMockContext({
        fluentDesignLanguage: {
          tokenTheme: {
            colorNeutralBackground1: "#custom",
            // everything else is missing
          },
        },
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.theme.cardBg).toBe("#custom");
      expect(element.props.theme.textPrimary).toBe(defaultTheme.textPrimary);
      expect(element.props.theme.border).toBe(defaultTheme.border);
    });

    it("returns defaultTheme when fluentDesignLanguage has no tokenTheme", () => {
      const ctx = createMockContext({
        fluentDesignLanguage: {},
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.theme).toEqual(defaultTheme);
    });
  });

  // ── detectRelatedFields (@syntax) ────

  describe("detectRelatedFields()", () => {
    it("detects @-prefixed slot values as related field references", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          subtitleField2: {
            raw: "@msdyn_serviceaccount",
            attributes: { LogicalName: "sub2", DisplayName: "Sub 2" },
          },
          tagField3: {
            raw: "@msdyn_priority",
            attributes: { LogicalName: "tag3", DisplayName: "Tag 3" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.relatedMappings).toHaveLength(2);
      expect(element.props.relatedMappings[0]).toEqual({
        sourceSlot: "titleField",
        fetchField: "msdyn_serviceaccount",
        targetSlot: "subtitleField2",
      });
      expect(element.props.relatedMappings[1]).toEqual({
        sourceSlot: "titleField",
        fetchField: "msdyn_priority",
        targetSlot: "tagField3",
      });
    });

    it("returns empty array when no slots have @-prefixed values", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);
      expect(element.props.relatedMappings).toEqual([]);
    });

    it("ignores slots whose raw is not a string", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          subtitleField1: {
            raw: 42,
            attributes: { LogicalName: "count", DisplayName: "Count" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);
      expect(element.props.relatedMappings).toEqual([]);
    });

    it("ignores strings not starting with @", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          detailField1: {
            raw: "some_regular_value",
            attributes: { LogicalName: "detail1", DisplayName: "Detail" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);
      expect(element.props.relatedMappings).toEqual([]);
    });

    it("ignores bare @ with empty field name", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          phoneField1: {
            raw: "@",
            attributes: { LogicalName: "phone1", DisplayName: "Phone" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);
      expect(element.props.relatedMappings).toEqual([]);
    });

    it("detects @. prefix as current-record mappings", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          gridField5: {
            raw: "@.resource.resourcetype",
            attributes: {},
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      // @. mappings go to currentRecordMappings, not relatedMappings
      expect(element.props.relatedMappings).toEqual([]);
      expect(element.props.currentRecordMappings).toHaveLength(1);
      expect(element.props.currentRecordMappings[0]).toEqual({
        sourceSlot: "__currentRecord__",
        fetchField: "resource.resourcetype",
        targetSlot: "gridField5",
      });
    });

    it("splits @ and @. mappings to separate props", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          subtitleField2: {
            raw: "@msdyn_serviceaccount",
            attributes: {},
          },
          gridField5: {
            raw: "@.resource.resourcetype",
            attributes: {},
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.relatedMappings).toHaveLength(1);
      expect(element.props.relatedMappings[0].sourceSlot).toBe("titleField");
      expect(element.props.currentRecordMappings).toHaveLength(1);
      expect(element.props.currentRecordMappings[0].sourceSlot).toBe("__currentRecord__");
    });
  });

  // ── formatLogicalName ─────────────────

  describe("formatLogicalName()", () => {
    function getFormatLogicalName(ctrl: InfoCard) {
      return asAny(ctrl).formatLogicalName.bind(ctrl) as (name: string) => string;
    }

    it("strips msdyn_ prefix and formats", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatLogicalName(control);
      expect(fmt("msdyn_workordertype")).toBe("Workordertype");
    });

    it("handles underscores as spaces", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatLogicalName(control);
      expect(fmt("booking_type")).toBe("Booking Type");
    });

    it("handles simple names", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatLogicalName(control);
      expect(fmt("starttime")).toBe("Starttime");
    });
  });

  // ── isDurationField (via readSlot) ────

  describe("isDurationField()", () => {
    function getIsDurationField(ctrl: InfoCard) {
      return asAny(ctrl).isDurationField.bind(ctrl) as (
        attrs: Record<string, unknown>,
        raw: unknown,
        formatted: string | undefined,
      ) => boolean;
    }

    it("returns true when Format attribute is Duration", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({ Format: "Duration" }, 60, undefined)).toBe(true);
    });

    it("returns true when Format attribute is duration (lowercase)", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({ Format: "duration" }, 60, undefined)).toBe(true);
    });

    it("returns false when raw is not a number", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({ Format: "Duration" }, "not-a-number", undefined)).toBe(false);
    });

    it("detects duration from formatted string ending in hours", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({}, 120, "2 hours")).toBe(true);
    });

    it("detects duration from formatted string ending in hour", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({}, 60, "1 hour")).toBe(true);
    });

    it("detects duration from formatted string ending in minutes", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({}, 30, "30 minutes")).toBe(true);
    });

    it("detects duration from formatted string ending in minute", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({}, 1, "1 minute")).toBe(true);
    });

    it("returns false for non-duration number field", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({}, 42, "42")).toBe(false);
    });

    it("returns false when no Format attr and no formatted string", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const isDuration = getIsDurationField(control);
      expect(isDuration({}, 100, undefined)).toBe(false);
    });
  });

  // ── formatDuration ────────────────────

  describe("formatDuration()", () => {
    function getFormatDuration(ctrl: InfoCard) {
      return asAny(ctrl).formatDuration.bind(ctrl) as (minutes: number) => string;
    }

    it("formats days, hours, and minutes", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatDuration(control);
      expect(fmt(1530)).toBe("1d 1h 30m");
    });

    it("formats hours and minutes only", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatDuration(control);
      expect(fmt(150)).toBe("2h 30m");
    });

    it("formats minutes only", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatDuration(control);
      expect(fmt(45)).toBe("45m");
    });

    it("formats hours only (no leftover minutes)", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatDuration(control);
      expect(fmt(120)).toBe("2h");
    });

    it("formats days only", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatDuration(control);
      expect(fmt(1440)).toBe("1d");
    });

    it('returns "0m" for zero', () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatDuration(control);
      expect(fmt(0)).toBe("0m");
    });

    it("returns string representation for negative numbers", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatDuration(control);
      expect(fmt(-10)).toBe("-10");
    });

    it("handles large values (multiple days)", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const fmt = getFormatDuration(control);
      // 3 days, 5 hours, 15 minutes = 4635 minutes
      expect(fmt(4635)).toBe("3d 5h 15m");
    });
  });

  // ── getOutputs ────────────────────────

  describe("getOutputs()", () => {
    it("returns empty object", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      const outputs = control.getOutputs();
      expect(outputs).toEqual({});
    });
  });

  // ── destroy ───────────────────────────

  describe("destroy()", () => {
    it("does not throw", () => {
      const ctx = createMockContext();
      control.init(ctx, notifyOutputChanged);
      expect(() => control.destroy()).not.toThrow();
    });
  });

  // ── contextInfo detection ────────────

  describe("contextInfo()", () => {
    it("stores form entity and record ID from contextInfo", () => {
      const ctx = createMockContext({
        contextInfo: {
          entityTypeName: "bookableresourcebooking",
          entityId: "bb111111-1111-1111-1111-111111111111",
        },
        slots: {
          titleField: {
            raw: "WO-123",
            attributes: { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.currentRecordEntityType).toBe("bookableresourcebooking");
      expect(element.props.currentRecordId).toBe("bb111111-1111-1111-1111-111111111111");
    });

    it("passes resolveRecordFields when contextInfo present and not design-time", () => {
      const ctx = createMockContext({
        contextInfo: {
          entityTypeName: "bookableresourcebooking",
          entityId: "bb111111-1111-1111-1111-111111111111",
        },
        slots: {
          titleField: {
            raw: "WO-123",
            attributes: { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.resolveRecordFields).toBeDefined();
    });

    it("omits resolveRecordFields at design time", () => {
      const ctx = createMockContext({
        contextInfo: {
          entityTypeName: "bookableresourcebooking",
          entityId: "bb111111-1111-1111-1111-111111111111",
        },
        slots: {
          // titleField with no raw = design time
          titleField: {
            raw: null,
            attributes: { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);

      expect(element.props.resolveRecordFields).toBeUndefined();
    });
  });

  // ── resolveRecordFieldsAsync ─────────

  describe("resolveRecordFieldsAsync()", () => {
    function getResolve(ctrl: InfoCard) {
      return asAny(ctrl).resolveRecordFieldsAsync.bind(ctrl) as (
        ctx: ComponentFramework.Context<IInputs>,
      ) => Promise<Record<string, { label: string; value: string; color?: string }>>;
    }

    it("matches numeric param to WebAPI column and resolves formatted value", async () => {
      const webAPIMock = jest.fn().mockResolvedValue(BOOKING_RECORD);
      const metadataMock = jest.fn().mockResolvedValue(makeMetadataResponse([
        { LogicalName: "bookingtype", DisplayName: "Booking Type", options: [
          { Value: 1, Label: "Solid" }, { Value: 2, Label: "Liquid" },
        ]},
        { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
      ]));

      const ctx = createMockContext({
        contextInfo: { entityTypeName: "bookableresourcebooking", entityId: "bb111111-1111-1111-1111-111111111111" },
        webAPIMock,
        metadataMock,
        slots: {
          titleField: {
            raw: [{ id: "ab111111-1111-1111-1111-111111111111", name: "WO-00047", entityType: "msdyn_workorder" }],
            attributes: { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
          },
          gridField3: { raw: 1, type: "OptionSet", attributes: {} },
        },
      });
      control.init(ctx, notifyOutputChanged);
      control.updateView(ctx); // sets _formEntityName + _formRecordId

      const overrides = await getResolve(control)(ctx);
      expect(overrides.gridField3).toBeDefined();
      expect(overrides.gridField3.label).toBe("Booking Type");
      expect(overrides.gridField3.value).toBe("Solid");
    });

    it("matches datetime param to WebAPI column", async () => {
      const startDate = new Date("2026-03-30T09:00:00Z");
      const webAPIMock = jest.fn().mockResolvedValue(makeWebAPIRecord({
        starttime: "2026-03-30T09:00:00Z",
        "starttime@OData.Community.Display.V1.FormattedValue": "3/30/2026 9:00 AM",
        statecode: 0, statuscode: 1,
      }));
      const metadataMock = jest.fn().mockResolvedValue(makeMetadataResponse([
        { LogicalName: "starttime", DisplayName: "Start Time" },
      ]));

      const ctx = createMockContext({
        contextInfo: { entityTypeName: "bookableresourcebooking", entityId: "bb111111-1111-1111-1111-111111111111" },
        webAPIMock,
        metadataMock,
        slots: {
          titleField: {
            raw: [{ id: "ab111111-1111-1111-1111-111111111111", name: "WO-00047", entityType: "msdyn_workorder" }],
            attributes: { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
          },
          gridField1: { raw: startDate, type: "DateAndTime.DateAndTime", attributes: {} },
        },
      });
      control.init(ctx, notifyOutputChanged);
      control.updateView(ctx);

      const overrides = await getResolve(control)(ctx);
      expect(overrides.gridField1).toBeDefined();
      expect(overrides.gridField1.label).toBe("Start Time");
      expect(overrides.gridField1.value).toBe("3/30/2026 9:00 AM");
    });

    it("skips system columns during value matching", async () => {
      const webAPIMock = jest.fn().mockResolvedValue(makeWebAPIRecord({
        statuscode: 1, statecode: 0, bookingtype: 1,
        "bookingtype@OData.Community.Display.V1.FormattedValue": "Solid",
      }));
      const metadataMock = jest.fn().mockResolvedValue(makeMetadataResponse([
        { LogicalName: "bookingtype", DisplayName: "Booking Type" },
      ]));

      const ctx = createMockContext({
        contextInfo: { entityTypeName: "bookableresourcebooking", entityId: "bb111111-1111-1111-1111-111111111111" },
        webAPIMock,
        metadataMock,
        slots: {
          titleField: {
            raw: [{ id: "ab111111-1111-1111-1111-111111111111", name: "WO", entityType: "msdyn_workorder" }],
            attributes: { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
          },
          // raw=1 matches both statuscode(1) and bookingtype(1)
          // System columns should be skipped → matches bookingtype
          gridField3: { raw: 1, type: "OptionSet", attributes: {} },
        },
      });
      control.init(ctx, notifyOutputChanged);
      control.updateView(ctx);

      const overrides = await getResolve(control)(ctx);
      expect(overrides.gridField3).toBeDefined();
      expect(overrides.gridField3.label).toBe("Booking Type");
    });

    it("resolves booking status color from lookup entity", async () => {
      const webAPIMock = jest.fn().mockImplementation((entityType: string) => {
        if (entityType === "bookableresourcebooking") {
          return Promise.resolve(makeWebAPIRecord({
            _bookingstatus_value: "bs111111-1111-1111-1111-111111111111",
          }));
        }
        if (entityType === "bookingstatus") {
          return Promise.resolve({ msdyn_statuscolor: "49F249" });
        }
        return Promise.resolve({});
      });

      const ctx = createMockContext({
        contextInfo: { entityTypeName: "bookableresourcebooking", entityId: "bb111111-1111-1111-1111-111111111111" },
        webAPIMock,
        slots: {
          titleField: {
            raw: [{ id: "ab111111-1111-1111-1111-111111111111", name: "WO", entityType: "msdyn_workorder" }],
            attributes: { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
          },
          tagField1: {
            raw: [{ id: "bs111111-1111-1111-1111-111111111111", name: "Scheduled", entityType: "bookingstatus" }],
            type: "Lookup.Simple",
            attributes: {},
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      control.updateView(ctx);

      const overrides = await getResolve(control)(ctx);
      expect(overrides.tagField1).toBeDefined();
      expect(overrides.tagField1.color).toBe("#49F249");
    });

    it("skips titleField in overrides", async () => {
      const webAPIMock = jest.fn().mockResolvedValue(BOOKING_RECORD);
      const metadataMock = jest.fn().mockResolvedValue(makeMetadataResponse([
        { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
      ]));

      const ctx = createMockContext({
        contextInfo: { entityTypeName: "bookableresourcebooking", entityId: "bb111111-1111-1111-1111-111111111111" },
        webAPIMock,
        metadataMock,
        slots: {
          titleField: {
            raw: [{ id: "ab111111-1111-1111-1111-111111111111", name: "WO-00047", entityType: "msdyn_workorder" }],
            attributes: { LogicalName: "msdyn_workorder", DisplayName: "Work Order" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      control.updateView(ctx);

      const overrides = await getResolve(control)(ctx);
      expect(overrides.titleField).toBeUndefined();
    });
  });

  // ── fetchRelatedFields ───────────────

  describe("fetchRelatedFields()", () => {
    function getFetch(ctrl: InfoCard) {
      return asAny(ctrl).fetchRelatedFields.bind(ctrl) as (
        ctx: ComponentFramework.Context<IInputs>,
        entityType: string, id: string, columns: string[], skipExpand?: boolean,
      ) => Promise<Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string; color?: string }>>;
    }

    it("reads direct field with OData formatted value", async () => {
      const woRecord = makeWebAPIRecord({
        msdyn_instructions: "Fix the boiler",
        "msdyn_instructions@OData.Community.Display.V1.FormattedValue": "Fix the boiler",
      });
      const webAPIMock = jest.fn().mockResolvedValue(woRecord);
      const metadataMock = jest.fn().mockResolvedValue(makeMetadataResponse([
        { LogicalName: "msdyn_instructions", DisplayName: "Instructions" },
      ]));

      const ctx = createMockContext({ webAPIMock, metadataMock });
      control.init(ctx, notifyOutputChanged);

      const results = await getFetch(control)(ctx, "msdyn_workorder", "wo-123", ["msdyn_instructions"]);
      expect(results.msdyn_instructions).toBeDefined();
      expect(results.msdyn_instructions.value).toBe("Fix the boiler");
      expect(results.msdyn_instructions.label).toBe("Instructions");
    });

    it("reads lookup field via _col_value pattern", async () => {
      const woRecord = makeWebAPIRecord({
        _msdyn_serviceaccount_value: "acct-123",
        "_msdyn_serviceaccount_value@OData.Community.Display.V1.FormattedValue": "Contoso Ltd",
        "_msdyn_serviceaccount_value@Microsoft.Dynamics.CRM.lookuplogicalname": "account",
      });
      const webAPIMock = jest.fn().mockResolvedValue(woRecord);
      const metadataMock = jest.fn().mockResolvedValue(makeMetadataResponse([
        { LogicalName: "msdyn_serviceaccount", DisplayName: "Service Account" },
      ]));

      const ctx = createMockContext({ webAPIMock, metadataMock });
      control.init(ctx, notifyOutputChanged);

      const results = await getFetch(control)(ctx, "msdyn_workorder", "wo-123", ["msdyn_serviceaccount"]);
      expect(results.msdyn_serviceaccount).toBeDefined();
      expect(results.msdyn_serviceaccount.value).toBe("Contoso Ltd");
      expect(results.msdyn_serviceaccount.lookupEntityType).toBe("account");
    });

    it("uses hop2 for @. dotted paths when skipExpand=true", async () => {
      const webAPIMock = jest.fn().mockImplementation(
        (entityType: string, _id: string, options: string) => {
          if (entityType === "bookableresourcebooking" && options.includes("_resource_value")) {
            return Promise.resolve(makeWebAPIRecord({
              _resource_value: "res-123",
              "_resource_value@OData.Community.Display.V1.FormattedValue": "John Smith",
              "_resource_value@Microsoft.Dynamics.CRM.lookuplogicalname": "bookableresource",
            }));
          }
          if (entityType === "bookableresource") {
            return Promise.resolve(makeWebAPIRecord({
              resourcetype: 2,
              "resourcetype@OData.Community.Display.V1.FormattedValue": "Contact",
            }));
          }
          return Promise.resolve({});
        },
      );
      const metadataMock = jest.fn().mockResolvedValue(makeMetadataResponse([
        { LogicalName: "resourcetype", DisplayName: "Resource Type" },
      ]));

      const ctx = createMockContext({ webAPIMock, metadataMock });
      control.init(ctx, notifyOutputChanged);

      const results = await getFetch(control)(
        ctx, "bookableresourcebooking", "bk-123", ["resource.resourcetype"], true,
      );
      expect(results["resource.resourcetype"]).toBeDefined();
      expect(results["resource.resourcetype"].value).toBe("Contact");
      expect(results["resource.resourcetype"].label).toBe("Resource Type");
    });

    it("returns empty object on complete failure", async () => {
      const webAPIMock = jest.fn().mockRejectedValue(new Error("Network error"));
      const ctx = createMockContext({ webAPIMock });
      control.init(ctx, notifyOutputChanged);

      const results = await getFetch(control)(ctx, "msdyn_workorder", "wo-123", ["msdyn_name"]);
      expect(results).toEqual({});
    });
  });

  // ── Realistic Dynamics runtime ───────

  describe("Dynamics runtime realism", () => {
    it("input slot has empty attributes by default", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
          gridField1: {
            raw: "some value",
          },
        },
      });
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const params = ctx.parameters as Record<string, any>;
      // titleField (bound) gets real attributes
      expect(params.titleField.attributes.LogicalName).toBe("name");
      // gridField1 (input) gets empty attributes like real Dynamics
      expect(params.gridField1.attributes).toEqual({});
    });

    it("readSlot returns field for input slot with empty attributes", () => {
      const ctx = createMockContext({
        slots: {
          gridField1: {
            raw: "test value",
            // No attributes override → gets {} (realistic input property)
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = asAny(control).readSlot.bind(control) as (
        ctx: ComponentFramework.Context<IInputs>, key: string,
      ) => unknown;
      const result = readSlot(ctx, "gridField1");
      // Should NOT be null — the field has data, it just lacks metadata
      expect(result).not.toBeNull();
    });
  });
});
