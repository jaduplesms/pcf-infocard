/**
 * Unit tests for the InfoCard PCF lifecycle class (index.ts).
 *
 * Covers: init, updateView, collectData, readSlot, readSlotGroup,
 *         resolveLayout, resolveTheme, parseRelatedConfig,
 *         isDurationField, formatDuration, getOutputs, destroy.
 */

import { InfoCard } from "../index";
import { IInputs } from "../generated/ManifestTypes";
import { defaultTheme } from "../InfoCard";

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
  slots?: Record<string, SlotParamOverride>;
}

function makeSlotParam(override?: SlotParamOverride) {
  if (!override) {
    return { raw: null, type: "Unknown", formatted: undefined, attributes: undefined };
  }
  return {
    raw: override.raw ?? null,
    type: override.type ?? "SingleLine.Text",
    formatted: override.formatted,
    attributes: override.attributes ?? { LogicalName: "field", DisplayName: "Field" },
  };
}

function createMockContext(overrides: ContextOverrides = {}): ComponentFramework.Context<IInputs> {
  const slots = overrides.slots ?? {};

  const allSlotKeys = [
    "titleField",
    "subtitleField1", "subtitleField2", "subtitleField3",
    "phoneField1", "phoneField2",
    "emailField", "webField",
    "latitudeField", "longitudeField",
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
    relatedConfig: { raw: overrides.relatedConfig ?? null },
  };

  for (const key of allSlotKeys) {
    parameters[key] = makeSlotParam(slots[key]);
  }

  const ctx: Record<string, unknown> = {
    parameters,
    mode: {
      isControlDisabled: false,
      allocatedWidth: 350,
      allocatedHeight: 500,
      isVisible: true,
      label: "InfoCard",
    } as ComponentFramework.Mode,
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
      getEntityMetadata: jest.fn().mockResolvedValue({
        Attributes: {
          getAll: () => [],
        },
      }),
    } as unknown as ComponentFramework.Utility,
    webAPI: {
      retrieveRecord: jest.fn().mockResolvedValue({}),
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

    it("returns null when attributes have no LogicalName", () => {
      const ctx = createMockContext({
        slots: {
          titleField: {
            raw: "test",
            attributes: {},
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const readSlot = getReadSlot(control);
      const result = readSlot(ctx, "titleField");
      expect(result).toBeNull();
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
      expect(readSlot(ctx2, "titleField")!.label).toBe("logname");
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

  // ── parseRelatedConfig ────────────────

  describe("parseRelatedConfig()", () => {
    it("parses sourceSlot:field>target format", () => {
      const ctx = createMockContext({
        relatedConfig: "titleField:msdyn_serviceaccount>subtitleField2,msdyn_priority>tagField3",
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
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

    it("returns empty array for null config", () => {
      const ctx = createMockContext({
        relatedConfig: null,
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

    it("returns empty array for empty string", () => {
      const ctx = createMockContext({
        relatedConfig: "",
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

    it("returns empty array when no colon separator", () => {
      const ctx = createMockContext({
        relatedConfig: "invalid-no-colon",
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

    it("skips entries without > separator", () => {
      const ctx = createMockContext({
        relatedConfig: "titleField:badentry,good_field>tagField1",
        slots: {
          titleField: {
            raw: "Title",
            attributes: { LogicalName: "name", DisplayName: "Name" },
          },
        },
      });
      control.init(ctx, notifyOutputChanged);
      const element = control.updateView(ctx);
      expect(element.props.relatedMappings).toHaveLength(1);
      expect(element.props.relatedMappings[0].fetchField).toBe("good_field");
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
});
