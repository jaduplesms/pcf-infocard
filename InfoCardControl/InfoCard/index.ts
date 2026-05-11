import { IInputs, IOutputs } from "./generated/ManifestTypes";
import {
    InfoCardComponent, InfoCardData, InfoCardTheme, SlotField, LayoutMode, defaultTheme,
    DEFAULT_STRINGS, formatLocalizedDuration,
} from "./InfoCard";
import type { RelatedFieldMapping, BindingDiagnostic, InfoCardStrings } from "./InfoCard";
import { isUnbindableColumn } from "./bindingValidator";
import * as React from "react";

// IMPORTANT: keep in sync with manifest version in ControlManifest.Input.xml.
// Bump both together on every deploy (mobile aggressively caches by manifest version).
const CONTROL_VERSION = "4.4.15";

/**
 * Column-name heuristic for Whole.Duration columns. Some Dataverse hosts don't
 * expose the IntegerAttributeMetadata.Format property under any documented key,
 * so metadata-based detection is unreliable. Falls back to recognizing common
 * column-name patterns: literal "duration", anything ending in "duration" or
 * "durationminutes" (e.g. msdyn_estimateddurationminutes, msdyn_actualduration).
 */
function isLikelyDurationColumn(colName: string | undefined): boolean {
    if (!colName) return false;
    const lc = colName.toLowerCase();
    return lc === "duration"
        || lc.endsWith("duration")
        || lc.endsWith("durationminutes");
}

// Minimal shape of context.formatting we use. The PCF typings don't expose
// formatInteger consistently across hosts, so we narrow it ourselves and
// guard every call.
interface MaybeFormatting {
    formatDateShort?: (d: Date, includeTime?: boolean) => string;
    formatTime?: (d: Date, behavior?: number) => string;
    formatInteger?: (n: number) => string;
}

/** Pull context.formatting if available; tolerate hosts (tests) that omit it. */
function getFormatting(context: ComponentFramework.Context<IInputs>): MaybeFormatting | undefined {
    return (context as unknown as { formatting?: MaybeFormatting }).formatting;
}

/**
 * Read a localized resx string with a safe fallback.
 *
 * The PCF runtime returns the KEY itself (not "" or undefined) when a key
 * isn't present in the active language's resx — so we treat `value === key`
 * as "missing" and substitute the English default.
 */
function readResxString(
    context: ComponentFramework.Context<IInputs>,
    key: string,
    fallback: string,
): string {
    try {
        const v = context.resources?.getString?.(key);
        return v && v !== key ? v : fallback;
    } catch {
        return fallback;
    }
}

/**
 * Build the localized strings bag from context.resources, falling back to
 * English defaults. This honors the user's Dataverse UI language when the
 * resx for that LCID is registered in ControlManifest.Input.xml.
 */
function getStrings(context: ComponentFramework.Context<IInputs>): InfoCardStrings {
    return {
        sectionContact: readResxString(context, "Section_Contact", DEFAULT_STRINGS.sectionContact),
        sectionDetails: readResxString(context, "Section_Details", DEFAULT_STRINGS.sectionDetails),
        sectionInfo: readResxString(context, "Section_Info", DEFAULT_STRINGS.sectionInfo),
        actionCall: readResxString(context, "Action_Call", DEFAULT_STRINGS.actionCall),
        actionEmail: readResxString(context, "Action_Email", DEFAULT_STRINGS.actionEmail),
        actionOpenInMaps: readResxString(context, "Action_OpenInMaps", DEFAULT_STRINGS.actionOpenInMaps),
        actionOpenWebsite: readResxString(context, "Action_OpenWebsite", DEFAULT_STRINGS.actionOpenWebsite),
        actionOpenRecord: readResxString(context, "Action_OpenRecord", DEFAULT_STRINGS.actionOpenRecord),
        cardExpand: readResxString(context, "Card_Expand", DEFAULT_STRINGS.cardExpand),
        cardCollapse: readResxString(context, "Card_Collapse", DEFAULT_STRINGS.cardCollapse),
        durationDaysSuffix: readResxString(context, "Duration_Days_Suffix", DEFAULT_STRINGS.durationDaysSuffix),
        durationHoursSuffix: readResxString(context, "Duration_Hours_Suffix", DEFAULT_STRINGS.durationHoursSuffix),
        durationMinutesSuffix: readResxString(context, "Duration_Minutes_Suffix", DEFAULT_STRINGS.durationMinutesSuffix),
        durationZero: readResxString(context, "Duration_Zero", DEFAULT_STRINGS.durationZero),
        actionCopy: readResxString(context, "Action_Copy", DEFAULT_STRINGS.actionCopy),
        actionCopied: readResxString(context, "Action_Copied", DEFAULT_STRINGS.actionCopied),
    };
}

// Slot group definitions — order matters for rendering
const SUBTITLE_KEYS = ["subtitleField1", "subtitleField2", "subtitleField3"] as const;
const PHONE_KEYS = ["phoneField1", "phoneField2"] as const;
const DETAIL_KEYS = ["detailField1", "detailField2", "detailField3"] as const;
const GRID_KEYS = [
    "gridField1", "gridField2", "gridField3", "gridField4", "gridField5", "gridField6",
] as const;
const TAG_KEYS = ["tagField1", "tagField2", "tagField3"] as const;
const ALL_SLOT_KEYS = [
    "titleField", "imageField", ...SUBTITLE_KEYS, ...PHONE_KEYS, "emailField", "webField",
    "addressField", "latitudeField", "longitudeField",
    ...DETAIL_KEYS, ...GRID_KEYS, ...TAG_KEYS,
];

// Opt-in slot-resolution tracing. Toggle in browser DevTools or harness:
//   (window as any).__INFOCARD_DEBUG__ = true;
// then re-trigger updateView. Off by default to keep mobile/host logs clean.
function isSlotDebugEnabled(): boolean {
    if (typeof window === "undefined") return false;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return (window as any).__INFOCARD_DEBUG__ === true;
}

// ────────────────────────────────────────
// Design-time / authoring mode sample data
// ────────────────────────────────────────
// Deterministic — no Date.now()/Math.random() so designer previews stay stable
// across re-renders and snapshot tests remain reproducible.
const SAMPLE_DATE = new Date(2026, 0, 15, 9, 30, 0, 0);

// Slot-name semantic overrides (preferred over type lookup).
// Uses the slot's intent (e.g. phone, email, address) rather than its underlying
// PCF type so the maker preview looks like a real card, not a row of "42"s.
// Design-time sample text. Deliberately reads as obvious placeholder copy (not realistic
// record data) so the maker can tell at a glance that nothing is bound to a real record yet.
// Each entry hints at the slot's purpose to help with authoring.
const SAMPLE_BY_SLOT: Record<string, string> = {
    titleField: "Title appears here",
    imageField: "",
    subtitleField1: "Subtitle 1",
    subtitleField2: "Subtitle 2",
    subtitleField3: "Subtitle 3",
    phoneField1: "Phone number 1",
    phoneField2: "Phone number 2",
    emailField: "Email address",
    webField: "Website URL",
    addressField: "Street address, city, region",
    latitudeField: "Latitude",
    longitudeField: "Longitude",
    detailField1: "Detail 1 will appear here — long-form text from the bound column such as instructions, description or summary.",
    detailField2: "Detail 2 will appear here — long-form text from the bound column.",
    detailField3: "Detail 3 will appear here — long-form text from the bound column.",
    tagField1: "Tag 1",
    tagField2: "Tag 2",
    tagField3: "Tag 3",
    // gridField1..6 fall through to SAMPLE_BY_TYPE so date/currency/number columns
    // render with type-appropriate sample values in the grid.
};

// Per-PCF-type fallback when slot has no semantic sample (e.g. gridField*).
// All values reuse obvious placeholder copy for the same reason as SAMPLE_BY_SLOT.
const SAMPLE_BY_TYPE: Record<string, { value: string; raw: unknown }> = {
    "SingleLine.Text": { value: "Sample text", raw: "Sample text" },
    "SingleLine.TextArea": { value: "Sample text — bound column value appears here.", raw: "Sample text — bound column value appears here." },
    "Multiple": { value: "Sample text — bound column value appears here.", raw: "Sample text — bound column value appears here." },
    "SingleLine.Email": { value: "name@example.com", raw: "name@example.com" },
    "SingleLine.Phone": { value: "+1 (555) 000-0000", raw: "+1 (555) 000-0000" },
    "SingleLine.URL": { value: "https://example.com", raw: "https://example.com" },
    "Whole.None": { value: "123", raw: 123 },
    "Currency": { value: "$0.00", raw: 0 },
    "Decimal": { value: "0.00", raw: 0 },
    "FP": { value: "0.00", raw: 0 },
    "DateAndTime.DateOnly": { value: "Jan 1, 2026", raw: SAMPLE_DATE },
    "DateAndTime.DateAndTime": { value: "Jan 1, 2026 09:00", raw: SAMPLE_DATE },
    "OptionSet": { value: "Choice", raw: "Choice" },
    "MultiSelectOptionSet": { value: "Choice 1, Choice 2", raw: "Choice 1, Choice 2" },
    "TwoOptions": { value: "Yes", raw: true },
    "Lookup.Simple": { value: "Lookup value", raw: "Lookup value" },
};

// Friendly names for slot keys in diagnostics
const SLOT_LABELS: Record<string, string> = {
    titleField: "Title", subtitleField1: "Subtitle 1", subtitleField2: "Subtitle 2", subtitleField3: "Subtitle 3",
    imageField: "Image / Avatar",
    phoneField1: "Phone 1", phoneField2: "Phone 2", emailField: "Email", webField: "Website",
    addressField: "Address", latitudeField: "Latitude", longitudeField: "Longitude",
    detailField1: "Detail 1", detailField2: "Detail 2", detailField3: "Detail 3",
    gridField1: "Grid 1", gridField2: "Grid 2", gridField3: "Grid 3",
    gridField4: "Grid 4", gridField5: "Grid 5", gridField6: "Grid 6",
    tagField1: "Tag 1", tagField2: "Tag 2", tagField3: "Tag 3",
};

// ────────────────────────────────────────
// Standard layout presets per table type
// ────────────────────────────────────────
// When a slot is not bound by the maker, the preset for the form entity (if known)
// fills it in by referencing a column on that same entity. The value is resolved via
// the existing resolveRecordFieldsAsync record fetch — no extra WebAPI calls.
//
// Maker bindings always win: a slot that has a `type` (i.e. is configured in form XML)
// is never overridden by a preset, even when the preset has an entry for that slot.
//
// Goal: drop the control on a known entity's form, bind only `titleField`, get a useful
// card for free. Makers can still override individual slots as needed.
type SlotPreset = Partial<Record<string, string>>;

const SLOT_PRESETS: Record<string, SlotPreset> = {
    msdyn_workorder: {
        subtitleField1: "msdyn_serviceaccount",
        subtitleField2: "msdyn_primaryincidenttype",
        addressField: "msdyn_displayaddress",
        detailField1: "msdyn_workordersummary",
        gridField1: "msdyn_systemstatus",
        gridField2: "msdyn_subtotalamount",
        tagField1: "msdyn_priority",
        tagField2: "msdyn_primaryincidenttype",
    },
    bookableresourcebooking: {
        subtitleField1: "resource",
        subtitleField2: "bookingstatus",
        gridField1: "starttime",
        gridField2: "endtime",
        gridField3: "duration",
        tagField1: "bookingstatus",
    },
    account: {
        subtitleField1: "primarycontactid",
        subtitleField2: "industrycode",
        imageField: "entityimage_url",
        phoneField1: "telephone1",
        phoneField2: "telephone2",
        emailField: "emailaddress1",
        webField: "websiteurl",
        addressField: "address1_composite",
        tagField1: "statuscode",
    },
    contact: {
        subtitleField1: "jobtitle",
        subtitleField2: "parentcustomerid",
        imageField: "entityimage_url",
        phoneField1: "mobilephone",
        phoneField2: "telephone1",
        emailField: "emailaddress1",
        webField: "websiteurl",
        addressField: "address1_composite",
        tagField1: "statuscode",
    },
    incident: {
        subtitleField1: "customerid",
        subtitleField2: "casetypecode",
        phoneField1: "customercontacted",
        emailField: "emailaddress",
        detailField1: "description",
        gridField1: "createdon",
        gridField2: "modifiedon",
        gridField3: "ticketnumber",
        gridField4: "casetypecode",
        tagField1: "prioritycode",
        tagField2: "statuscode",
        tagField3: "severitycode",
    },
    msdyn_customerasset: {
        subtitleField1: "msdyn_account",
        subtitleField2: "msdyn_product",
        addressField: "msdyn_address1_composite",
        latitudeField: "msdyn_latitude",
        longitudeField: "msdyn_longitude",
        detailField1: "msdyn_description",
        gridField1: "msdyn_serialnumber",
        gridField2: "msdyn_registrationnumber",
        gridField3: "msdyn_installdate",
        gridField4: "msdyn_warrantystartdate",
        gridField5: "msdyn_warrantyenddate",
        tagField1: "statuscode",
    },
    msdyn_workorderservicetask: {
        subtitleField1: "msdyn_workorder",
        subtitleField2: "msdyn_tasktype",
        detailField1: "msdyn_description",
        gridField1: "msdyn_estimateddurationminutes",
        gridField2: "msdyn_actualdurationminutes",
        gridField3: "msdyn_lineorder",
        gridField4: "msdyn_percentcomplete",
        tagField1: "msdyn_iscompleted",
    },
    msdyn_agreement: {
        subtitleField1: "msdyn_serviceaccount",
        subtitleField2: "msdyn_billingaccount",
        addressField: "msdyn_serviceaddress_composite",
        detailField1: "msdyn_description",
        gridField1: "msdyn_startdate",
        gridField2: "msdyn_enddate",
        gridField3: "msdyn_systemstatus",
        gridField4: "msdyn_substatus",
        tagField1: "msdyn_systemstatus",
    },
};

interface ColumnMeta {
    displayName: string;
    logicalName: string;
    /** Format hint from metadata (e.g. "duration", "phone", "email"). Lowercased. */
    format?: string;
    options?: Array<{ value: number; label: string; color?: string }>;
}

export class InfoCard implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private context: ComponentFramework.Context<IInputs>;

    // Stable references for React useEffect callbacks
    private readonly boundFetchRelatedData = (entityType: string, id: string, columns: string[]) => {
        const isCurrentRecord = entityType === this._formEntityName && id === this._formRecordId;
        return this.fetchRelatedFields(this.context, entityType, id, columns, isCurrentRecord);
    };
    private readonly boundResolveRecordFields = () => {
        return this.resolveRecordFieldsAsync(this.context);
    };

    // Form entity context
    private _formEntityName: string | null = null;
    private _formRecordId: string | null = null;
    /** Cached per-updateView; consumed by readSlot for duration formatting. */
    private _formatting?: MaybeFormatting;
    private _strings: InfoCardStrings = DEFAULT_STRINGS;

    // Entity metadata cache for bound field labels + option set resolution
    private _columnMetadata: Record<string, ColumnMeta> = {};
    private _slotToColumn: Record<string, string> = {};
    private _resolvedValues: Record<string, string> = {};
    // Tracks slots whose column mapping originated from SLOT_PRESETS (vs. maker
    // binding via attributes.LogicalName, vs. value-matching at runtime). Used by
    // the override-merger to decide whether to surface the real column DisplayName
    // ("Start Time") or fall back to a generic slot label ("Grid 1") — value-matched
    // slots get generic labels so the row labels stay consistent across slots that
    // happen to match vs. ones that don't.
    private _presetSlots: Set<string> = new Set();
    // Colors resolved from lookup entity records
    private _resolvedColors: Record<string, string> = {};
    // True when rendering inside the Power Apps form designer. When set, readSlot
    // substitutes deterministic sample data and updateView suppresses WebAPI fetches.
    private _isAuthoringMode = false;

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
    ): void {
        this.context = context;
        console.log(`[InfoCard] v${CONTROL_VERSION} initialized`);
    }

    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        this.context = context;
        // Detect form-designer (authoring) mode.
        // Primary: undocumented context.mode.isAuthoringMode flag (true ONLY in designer).
        // Fallback 1: location.ancestorOrigins[0] === make.powerapps.com — resilient if MS
        //   ever renames or moves the flag.
        // Fallback 2 (test harness): window.__INFOCARD_AUTHORING__ === true OR location.hash
        //   contains "authoring" OR ?authoring=1 — lets `npm start` preview the designer path
        //   locally without needing a real Power Apps host. See README "Authoring-mode preview".
        // Refs: itmustbecode.com (2025-06), butenko.pro (2023-01).
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const modeRecord = context.mode as unknown as Record<string, any>;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const harnessFlag = typeof window !== "undefined" && (window as any).__INFOCARD_AUTHORING__ === true;
        const urlFlag = typeof location !== "undefined" && (
            (location.hash || "").toLowerCase().includes("authoring") ||
            (location.search || "").toLowerCase().includes("authoring=1")
        );
        this._isAuthoringMode = modeRecord.isAuthoringMode === true ||
            (typeof location !== "undefined" && (
                location.ancestorOrigins?.[0] === "https://make.powerapps.com" ||
                location.ancestorOrigins?.[0] === "https://make.preview.powerapps.com"
            )) ||
            harnessFlag ||
            urlFlag;

        // Detect form entity context for metadata resolution and @. syntax.
        // PCF instances can be reused across SPA navigations between records — re-read on every
        // updateView and clear record-scoped caches when entity or recordId changes, otherwise
        // resolveRecordFieldsAsync would bind to a stale GUID and overrides would be wrong.
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const modeAny = context.mode as Record<string, any>;
        const newEntity: string | null = modeAny.contextInfo?.entityTypeName ?? null;
        const newRecordId: string | null = modeAny.contextInfo?.entityId
            ? this.formatGuid(String(modeAny.contextInfo.entityId))
            : null;
        if (newEntity && (newEntity !== this._formEntityName || newRecordId !== this._formRecordId)) {
            if (this._formEntityName !== null) {
                // Record context changed — clear caches scoped to the previous record.
                this._slotToColumn = {};
                this._resolvedValues = {};
                this._resolvedColors = {};
                // Column metadata is keyed by logical name and is per-entity, so only clear it
                // when the entity itself changed.
                if (newEntity !== this._formEntityName) {
                    this._columnMetadata = {};
                }
            }
            this._formEntityName = newEntity;
            this._formRecordId = newRecordId;
        }

        const strings = getStrings(context);
        const formatting = getFormatting(context);
        this._strings = strings;
        this._formatting = formatting;

        const data = this.collectData(context, formatting, strings);
        const layout = this.resolveLayout(context);
        const hideEmpty = context.parameters.hideEmptyFields?.raw !== false; // default true
        const showBorder = context.parameters.showCardBorder?.raw !== false; // default true
        const showVersion = context.parameters.showVersionInfo?.raw === true; // default false
        const showTitle = context.parameters.showTitle?.raw !== false; // default true
        const startExpanded = context.parameters.startExpanded?.raw !== false; // default true
        const subtitleSeparator = context.parameters.subtitleSeparator?.raw || undefined;
        const titlePrefix = context.parameters.titlePrefix?.raw || undefined;
        const imageShape = (context.parameters.imageShape?.raw as "rounded" | "circle" | "square" | null) || "rounded";
        const collapsibleSections = (context.parameters.collapsibleSections?.raw as "none" | "body" | "body-tags" | "all" | null) || "body";
        const showDetailIcons = context.parameters.showDetailIcons?.raw !== false; // default true
        const detailLabelStyle = (context.parameters.detailLabelStyle?.raw as "none" | "inline-bold" | "above" | null) || "none";
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const clientAny = (context as any).client;
        const formFactor: number | undefined = (() => {
            try {
                const v = clientAny?.getFormFactor?.();
                return typeof v === "number" ? v : undefined;
            } catch { return undefined; }
        })();
        const theme = this.resolveTheme(context);

        // Detect all related field mappings, split by source
        const allMappings = this.detectRelatedFields(context);
        // If titleField is a lookup whose target entity has a SLOT_PRESETS entry,
        // synthesize @-prefixed mappings for slots the maker didn't bind so the
        // existing related-field pipeline auto-fills them from the lookup record.
        const titleLookupMappings = this.applyTitleLookupPreset(context, data, allMappings);
        const titleMappings = [...allMappings.filter(m => m.sourceSlot === "titleField"), ...titleLookupMappings];
        const currentRecordMappings = allMappings.filter(m => m.sourceSlot === "__currentRecord__");

        // Design-time detection: titleField is configured (valid type) but has no record data
        const titleParam = context.parameters.titleField;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const titleType = (titleParam as Record<string, any>)?.type;
        const designTime = titleParam != null
            && titleType !== "Unknown"
            && (titleParam.raw === null || titleParam.raw === undefined);

        // Build binding diagnostics for design-time panel (only warnings)
        const bindingDiagnostics = designTime ? this.buildBindingDiagnostics(context, allMappings) : undefined;

        const onOpenRecord = (entityType: string, id: string) => {
            context.navigation.openForm({
                entityName: entityType,
                entityId: id,
            });
        };

        // Record resolution is triggered by React useEffect via resolveRecordFields callback

        console.log("[InfoCard] title lookup:", data.title?.lookupEntityType, data.title?.lookupId,
            "| title mappings:", titleMappings.length,
            "| current-record mappings:", currentRecordMappings.length);

        return React.createElement(InfoCardComponent, {
            data,
            layout,
            hideEmpty,
            showBorder,
            showVersion,
            showTitle,
            startExpanded,
            subtitleSeparator,
            titlePrefix,
            imageShape,
            collapsibleSections,
            showDetailIcons,
            detailLabelStyle,
            formFactor,
            designTime,
            theme,
            version: CONTROL_VERSION,
            // Suppress all fetch pathways while authoring — sample data is already
            // baked into `data` by readSlot, and the maker isn't bound to a real record.
            relatedMappings: this._isAuthoringMode ? [] : titleMappings,
            currentRecordMappings: this._isAuthoringMode ? [] : currentRecordMappings,
            currentRecordEntityType: this._formEntityName ?? undefined,
            currentRecordId: this._formRecordId ?? undefined,
            fetchRelatedData: this._isAuthoringMode ? undefined : this.boundFetchRelatedData,
            resolveRecordFields: (!designTime && !this._isAuthoringMode && this._formEntityName && this._formRecordId)
                ? this.boundResolveRecordFields : undefined,
            onOpenRecord,
            bindingDiagnostics,
            strings,
        });
    }

    public getOutputs(): IOutputs {
        return {};
    }

    public destroy(): void { }

    // ────────────────────────────────────────
    // Data collection
    // ────────────────────────────────────────

    private collectData(
        context: ComponentFramework.Context<IInputs>,
        formatting?: MaybeFormatting,
        strings?: InfoCardStrings,
    ): InfoCardData {
        // Pre-register preset slot→column mappings BEFORE readSlot runs so its label
        // chain picks up the preset column name (e.g. msdyn_serviceaccount → "Service
        // Account") when no maker binding/displayName is available. At runtime these
        // mappings are also used by resolveRecordFieldsAsync to fetch real values.
        this.registerPresetMappings(context);

        // Read lat/lng as numbers
        const latSlot = this.readSlot(context, "latitudeField");
        const lngSlot = this.readSlot(context, "longitudeField");
        const lat = latSlot && !latSlot.isEmpty ? Number(latSlot.rawValue) : null;
        const lng = lngSlot && !lngSlot.isEmpty ? Number(lngSlot.rawValue) : null;

        const imageSlot = this.readSlot(context, "imageField");
        const imageUrl = (imageSlot && !imageSlot.isEmpty && typeof imageSlot.rawValue === "string" && imageSlot.rawValue.length > 0)
            ? imageSlot.rawValue
            : null;

        const data: InfoCardData = {
            title: this.readSlot(context, "titleField"),
            subtitles: this.readSlotGroup(context, SUBTITLE_KEYS),
            phones: this.readSlotGroup(context, PHONE_KEYS),
            email: this.readSlot(context, "emailField"),
            web: this.readSlot(context, "webField"),
            address: this.readSlot(context, "addressField"),
            latitude: (lat != null && !isNaN(lat)) ? lat : null,
            longitude: (lng != null && !isNaN(lng)) ? lng : null,
            details: this.readSlotGroup(context, DETAIL_KEYS),
            gridFields: this.readSlotGroup(context, GRID_KEYS),
            tags: this.readSlotGroup(context, TAG_KEYS),
            imageUrl,
        };

        // Apply standard layout preset for the form entity (if known) to slots the
        // maker did not bind. Inserts placeholders into the data and registers the
        // slot→column mapping so resolveRecordFieldsAsync resolves the value via the
        // record fetch already in flight. In authoring mode the placeholder carries
        // a sample value + preset-derived label so the maker sees a populated preview.
        this.applyPreset(context, data);

        return data;
    }

    /**
     * Register slot→column mappings from SLOT_PRESETS for the current form entity.
     * Must run before readSlot so the label resolution chain in readSlot picks up the
     * preset's column name (e.g. msdyn_serviceaccount → "Service Account") when no
     * maker binding is set. Only registers slots the maker has not configured.
     */
    private registerPresetMappings(
        context: ComponentFramework.Context<IInputs>,
    ): void {
        // Skip presets entirely in the form designer — sample data injection in
        // readSlot already covers unconfigured slots, and presets would cause
        // generic-labelled slots to render as if column-bound (confusing the maker
        // about what is actually configured vs. inferred).
        if (this._isAuthoringMode) return;
        const formEntity = this._formEntityName;
        if (!formEntity) return;
        const preset = SLOT_PRESETS[formEntity];
        if (!preset) return;

        for (const [slotKey, columnName] of Object.entries(preset)) {
            if (!columnName) continue;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const param = (context.parameters as Record<string, any>)[slotKey];
            const isMakerConfigured = !!(param && param.type && param.type !== "Unknown");
            if (isMakerConfigured) continue;
            if (this._slotToColumn[slotKey]) continue;
            this._slotToColumn[slotKey] = columnName;
            this._presetSlots.add(slotKey);
        }
    }

    /**
     * Fill unconfigured slots from SLOT_PRESETS for the current form entity. A slot
     * is considered unconfigured if the maker did not set a `type` on the property
     * (i.e. didn't bind a column or static value in the form designer). Configured
     * slots are left untouched.
     *
     * The slot→column mapping is recorded in `_slotToColumn` so the existing
     * resolveRecordFieldsAsync flow picks up labels and formatted values from the
     * already-fetched form record. A placeholder SlotField is injected into the data
     * so the slot has a render position; the override pass populates the actual value.
     */
    private applyPreset(
        context: ComponentFramework.Context<IInputs>,
        data: InfoCardData,
    ): void {
        const formEntity = this._formEntityName;
        if (!formEntity) return;
        const preset = SLOT_PRESETS[formEntity];
        if (!preset) return;

        for (const [slotKey, columnName] of Object.entries(preset)) {
            if (!columnName) continue;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const param = (context.parameters as Record<string, any>)[slotKey];
            const isMakerConfigured = !!(param && param.type && param.type !== "Unknown");
            if (isMakerConfigured) continue;

            // Reserve the slot→column mapping for resolveRecordFieldsAsync. Marker so
            // we know this slot was preset-driven and should be hidden when empty even
            // when hideEmpty=false (preset slots are speculative; don't show "---" rows
            // for columns that turn out to have no data on this record).
            this._slotToColumn[slotKey] = columnName;
            this._presetSlots.add(slotKey);

            // In authoring mode the maker has no record bound, so the override pass
            // never runs. Build a fully-rendered preview placeholder using the preset
            // column name as the label and a slot-appropriate sample as the value, so
            // the designer surfaces what the preset will populate at runtime when only
            // the title is bound.
            const presetLabel = this.formatLogicalName(columnName);
            const placeholder: SlotField = this._isAuthoringMode
                ? {
                    slotKey,
                    label: presetLabel,
                    value: SAMPLE_BY_SLOT[slotKey] ?? presetLabel,
                    rawValue: SAMPLE_BY_SLOT[slotKey] ?? presetLabel,
                    isEmpty: false,
                    isPreset: true,
                }
                : {
                    slotKey,
                    // Generic label here; the override pass will swap in the real
                    // DisplayName ("Start Time") once entity metadata loads.
                    label: this.formatPropertyKey(slotKey),
                    value: "---",
                    rawValue: null,
                    isEmpty: true,
                    isPreset: true,
                };

            if (slotKey.startsWith("subtitleField")) {
                if (!data.subtitles.some(s => s.slotKey === slotKey)) data.subtitles.push(placeholder);
            } else if (slotKey.startsWith("phoneField")) {
                if (!data.phones.some(s => s.slotKey === slotKey)) data.phones.push(placeholder);
            } else if (slotKey.startsWith("detailField")) {
                if (!data.details.some(s => s.slotKey === slotKey)) data.details.push(placeholder);
            } else if (slotKey.startsWith("gridField")) {
                if (!data.gridFields.some(s => s.slotKey === slotKey)) data.gridFields.push(placeholder);
            } else if (slotKey.startsWith("tagField")) {
                if (!data.tags.some(s => s.slotKey === slotKey)) data.tags.push(placeholder);
            } else if (slotKey === "emailField" && !data.email) {
                data.email = placeholder;
            } else if (slotKey === "webField" && !data.web) {
                data.web = placeholder;
            } else if (slotKey === "addressField" && !data.address) {
                data.address = placeholder;
            }
        }
    }

    /**
     * When titleField is a lookup AND the lookup's target entity has a SLOT_PRESETS
     * entry, synthesize @-prefixed related-field mappings for any slot the maker did
     * not bind and that the host preset did not already claim. Lets a control on a
     * Booking form with only titleField=msdyn_workorder render a full Work Order
     * summary card without explicit @msdyn_serviceaccount/@msdyn_address1 bindings.
     *
     * Returns the synthesized mappings to be appended to `titleMappings` so they're
     * fetched and merged through the existing related-field pipeline.
     */
    private applyTitleLookupPreset(
        context: ComponentFramework.Context<IInputs>,
        data: InfoCardData,
        existingMappings: RelatedFieldMapping[],
    ): RelatedFieldMapping[] {
        if (this._isAuthoringMode) return [];
        const lookupEntity = data.title?.lookupEntityType;
        if (!lookupEntity) return [];
        const preset = SLOT_PRESETS[lookupEntity];
        if (!preset) return [];

        const synthesized: RelatedFieldMapping[] = [];
        const claimedTargets = new Set(existingMappings.map(m => m.targetSlot));

        for (const [slotKey, columnName] of Object.entries(preset)) {
            if (!columnName) continue;
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const param = (context.parameters as Record<string, any>)[slotKey];
            const isMakerConfigured = !!(param && param.type && param.type !== "Unknown");
            if (isMakerConfigured) continue;
            if (this._slotToColumn[slotKey]) continue; // host preset already claimed it
            if (claimedTargets.has(slotKey)) continue;  // explicit @-mapping already targets it
            if (slotKey === "titleField") continue;     // never overwrite the title itself

            this._presetSlots.add(slotKey);

            // Placeholder so the slot has a render position before the fetch resolves.
            // Real label + value land via mergeRelatedFields once fetchRelatedFields returns.
            const placeholder: SlotField = {
                slotKey,
                label: this.formatPropertyKey(slotKey),
                value: "---",
                rawValue: null,
                isEmpty: true,
                isPreset: true,
            };

            if (slotKey.startsWith("subtitleField")) {
                if (!data.subtitles.some(s => s.slotKey === slotKey)) data.subtitles.push(placeholder);
            } else if (slotKey.startsWith("phoneField")) {
                if (!data.phones.some(s => s.slotKey === slotKey)) data.phones.push(placeholder);
            } else if (slotKey.startsWith("detailField")) {
                if (!data.details.some(s => s.slotKey === slotKey)) data.details.push(placeholder);
            } else if (slotKey.startsWith("gridField")) {
                if (!data.gridFields.some(s => s.slotKey === slotKey)) data.gridFields.push(placeholder);
            } else if (slotKey.startsWith("tagField")) {
                if (!data.tags.some(s => s.slotKey === slotKey)) data.tags.push(placeholder);
            } else if (slotKey === "emailField" && !data.email) {
                data.email = placeholder;
            } else if (slotKey === "webField" && !data.web) {
                data.web = placeholder;
            } else if (slotKey === "addressField" && !data.address) {
                data.address = placeholder;
            }

            synthesized.push({
                sourceSlot: "titleField",
                fetchField: columnName,
                targetSlot: slotKey,
            });
        }

        if (synthesized.length > 0) {
            console.log(`[InfoCard] Title-lookup preset (${lookupEntity}) added ${synthesized.length} mappings:`,
                synthesized.map(m => `${m.targetSlot}→${m.fetchField}`).join(", "));
        }
        return synthesized;
    }

    private readSlotGroup(
        context: ComponentFramework.Context<IInputs>,
        keys: readonly string[],
    ): SlotField[] {
        const fields: SlotField[] = [];
        for (const key of keys) {
            const field = this.readSlot(context, key);
            if (field) {
                fields.push(field);
            }
        }
        return fields;
    }

    private readSlot(
        context: ComponentFramework.Context<IInputs>,
        key: string,
    ): SlotField | null {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const param = (context.parameters as Record<string, any>)[key];
        if (!param) {
            if (isSlotDebugEnabled()) console.log("[InfoCard.readSlot]", key, "→ null (no param)");
            return null;
        }

        // Skip unconfigured properties (no type, no data, and no attributes)
        const hasAttributes = param.attributes && (param.attributes.DisplayName || param.attributes.LogicalName
            || param.attributes.displayName || param.attributes.logicalName);
        // type === "Unknown" is the explicit "this slot is not configured" signal — always honor it
        if (param.type === "Unknown") {
            if (isSlotDebugEnabled()) console.log("[InfoCard.readSlot]", key, "→ null (type=Unknown)");
            return null;
        }

        // Skip unconfigured properties (no type, no data, and no attributes)
        // EXCEPTION: in authoring mode (form designer), param.type is `undefined` and
        // attributes are empty for declared-but-unbound slots — we still want to render
        // sample data for them so the maker sees a populated preview. Skip these guards
        // entirely under authoring mode and let getDesignTimeSample drive the output.
        if (!this._isAuthoringMode) {
            if (!param.type && !hasAttributes && (param.raw === null || param.raw === undefined)) {
                if (isSlotDebugEnabled()) console.log("[InfoCard.readSlot]", key, "→ null (unconfigured)");
                return null;
            }
            // Attributes exist but have no identifying metadata and no type — field is not properly configured
            // Static input properties (static="true" in form XML) may have attributes={} but valid type and data
            if (param.attributes && !hasAttributes && !param.type) {
                if (isSlotDebugEnabled()) console.log("[InfoCard.readSlot]", key, "→ null (attrs without metadata, no type)");
                return null;
            }
        }

        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const attrs = (param as Record<string, any>).attributes;
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const paramAny = param as Record<string, any>;

        // ── Extract logical name ──
        const logicalName: string | null = attrs?.LogicalName ?? attrs?.logicalName ?? null;

        // Track slot → column mapping for entity metadata fetch
        if (logicalName) {
            this._slotToColumn[key] = logicalName;
        }

        // ── Extract display name from multiple sources ──
        let displayName: string | null = null;

        // 1. From entity metadata cache (most reliable, populated async)
        if (logicalName && this._columnMetadata[logicalName]) {
            displayName = this._columnMetadata[logicalName].displayName;
        }

        // 2. From param.attributes.DisplayName (various formats)
        if (!displayName) {
            displayName = this.extractDisplayNameFromMeta(attrs?.DisplayName ?? attrs?.displayName);
        }

        // 3. Param-level display name (some PCF hosts)
        if (!displayName && typeof paramAny.DisplayName === "string" && paramAny.DisplayName.length > 0) {
            displayName = paramAny.DisplayName;
        }

        // Check if we resolved the column name from the record fetch
        const resolvedColumn = this._slotToColumn[key] ?? logicalName;

        const raw = param.raw;

        // Label fallback chain: displayName (from attrs) → metadata cache → formatted
        // logicalName from attrs → generic slot key. We deliberately do NOT use
        // _slotToColumn[key] for the formatLogicalName fallback: that map is also
        // populated by SLOT_PRESETS and runtime value-matching, and surfacing those
        // column names as labels causes inconsistency (some slots show real column
        // DisplayNames, others show generic "Grid 1"). Preset/value-matched slots
        // get their real DisplayName via the metadata cache once it loads, or fall
        // through to the generic formatPropertyKey label.
        const cachedMeta = resolvedColumn ? this._columnMetadata[resolvedColumn] : null;
        const label = displayName
            ?? cachedMeta?.displayName
            ?? (logicalName ? this.formatLogicalName(logicalName) : null)
            ?? this.formatPropertyKey(key);

        const formatted = param.formatted;
        const resolvedValue = this._resolvedValues[key]; // from WebAPI record fetch

        // Authoring-mode (form-designer) sample injection.
        // Substitutes deterministic sample data when the slot is configured but the
        // designer can't supply a real value (raw=null, or @-prefix related-field
        // placeholder that won't resolve without a real record). Returns early to
        // bypass the rest of the formatting pipeline; no lookupId/lookupEntityType
        // is set so useEffect won't fire WebAPI requests.
        // NOTE: param.type is often `undefined` (not "Unknown") in the real form designer
        // even for declared slots, so we don't gate on it here — getDesignTimeSample falls
        // back to slot-name-based samples first and only consults type as a secondary lookup.
        if (this._isAuthoringMode && param.type !== "Unknown") {
            const isPlaceholder = raw === null || raw === undefined ||
                (typeof raw === "string" && raw.startsWith("@"));
            if (isPlaceholder) {
                const sample = this.getDesignTimeSample(key, String(param.type ?? ""));
                if (sample) {
                    return {
                        slotKey: key,
                        label,
                        value: sample.value,
                        rawValue: sample.raw,
                        isEmpty: false,
                    };
                }
            }
        }

        let displayValue = "";
        let isEmpty = false;

        // @-prefixed values are related field placeholders — treat as empty
        // The actual value will be populated by the related field fetch
        if (typeof raw === "string" && raw.startsWith("@")) {
            isEmpty = true;
            // Show the raw reference as placeholder for developer visibility
            displayValue = raw;
        } else if (raw === null || raw === undefined) {
            isEmpty = true;
            displayValue = "---";
        } else if (resolvedValue) {
            // Use the formatted value from the WebAPI record (includes OptionSet labels, dates, etc.)
            displayValue = resolvedValue;
        } else if (formatted) {
            displayValue = formatted;
        } else if (raw instanceof Date) {
            // Date objects (DateTime fields when formatted is null)
            displayValue = raw.toLocaleDateString(undefined, {
                year: "numeric", month: "short", day: "numeric",
                hour: "2-digit", minute: "2-digit",
            });
        } else if (typeof raw === "object" && Array.isArray(raw) && raw.length > 0) {
            // Lookup
            displayValue = raw[0].name ?? String(raw[0].id ?? "---");
        } else if (typeof raw === "object" && !Array.isArray(raw)) {
            // Single EntityReference object: {Id: {_rawGuid, _formattedGuid}, Name, LogicalName}
            // or simple lookup object: {name, id, entityType}
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const obj = raw as Record<string, any>;
            displayValue = obj.name ?? obj.Name ?? "";
        } else {
            displayValue = String(raw);
        }

        if (displayValue === "" || displayValue === "---") {
            isEmpty = true;
        }

        // OptionSet: if displayValue is a raw number, resolve from metadata cache
        if (typeof raw === "number" && !formatted && !resolvedValue && resolvedColumn) {
            const cached = this._columnMetadata[resolvedColumn];
            if (cached?.options) {
                const opt = cached.options.find(o => o.value === raw);
                if (opt) {
                    displayValue = opt.label;
                    isEmpty = false;
                }
            }
        }

        // Format duration fields (stored as minutes in Dataverse)
        if (!isEmpty && attrs && this.isDurationField(attrs, raw, formatted)) {
            displayValue = this.formatDuration(Number(raw), this._formatting, this._strings);
        }

        // Extract lookup entity/id for navigation and related field fetches
        const lookupRef = this.extractLookupRef(raw);
        let lookupEntityType = lookupRef?.entityType;
        const lookupId = lookupRef?.id;

        // Fallback: PCF may omit entityType — resolve from field metadata Targets
        if (lookupId && !lookupEntityType && attrs?.Targets) {
            const targets = Array.isArray(attrs.Targets) ? attrs.Targets
                : typeof attrs.Targets.getAll === "function" ? attrs.Targets.getAll() : [];
            if (targets.length > 0) lookupEntityType = targets[0];
        }

        // Resolve color: OptionSet from metadata cache, or lookup from resolved colors
        let optionColor: string | undefined = this._resolvedColors[key];
        if (!optionColor && typeof raw === "number" && resolvedColumn) {
            const cached = this._columnMetadata[resolvedColumn];
            const opt = cached?.options?.find(o => o.value === raw);
            if (opt?.color) optionColor = opt.color;
        }

        // Pre-compute split date/time text for DateAndTime.DateAndTime fields so the renderer
        // can stack them on two lines. Honors user locale formats via context.formatting.
        // DateOnly fields skip timeText. Falls back to toLocaleDateString/toLocaleTimeString
        // when context.formatting is unavailable (test harness).
        let dateText: string | undefined;
        let timeText: string | undefined;
        if (!isEmpty && raw instanceof Date) {
            const paramTypeStr = String(param.type ?? "");
            const isDateOnly = paramTypeStr === "DateAndTime.DateOnly";
            const isDateAndTime = paramTypeStr === "DateAndTime.DateAndTime";
            if (isDateAndTime || isDateOnly) {
                const fmt = this._formatting ?? getFormatting(this.context);
                try {
                    dateText = fmt?.formatDateShort
                        ? fmt.formatDateShort(raw, false)
                        : raw.toLocaleDateString();
                } catch {
                    dateText = raw.toLocaleDateString();
                }
                if (isDateAndTime) {
                    // Don't use context.formatting.formatTime here: on some Dataverse
                    // hosts it returns the full date+time string regardless of the
                    // behavior arg, which produces a duplicated date row in the UI.
                    // Intl.DateTimeFormat is consistent across hosts.
                    //
                    // Honor the user's Dataverse short-time pattern for 12/24h:
                    // shortTimePattern uses .NET tokens — "tt" indicates an AM/PM
                    // designator (12-hour); uppercase "HH" indicates 24-hour.
                    // Without this, browsers in en-GB/de-DE locales render 14:30
                    // even when the Dataverse user setting is "h:mm tt".
                    const us = (this.context as unknown as { userSettings?: { dateFormattingInfo?: { shortTimePattern?: string } } }).userSettings;
                    const stp = us?.dateFormattingInfo?.shortTimePattern ?? "";
                    let hour12: boolean | undefined = undefined;
                    if (stp.includes("tt")) hour12 = true;
                    else if (/\bH+\b/.test(stp)) hour12 = false;
                    try {
                        timeText = new Intl.DateTimeFormat(undefined, {
                            hour: "numeric", minute: "2-digit",
                            ...(hour12 !== undefined ? { hour12 } : {}),
                        }).format(raw);
                    } catch {
                        timeText = raw.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
                    }
                }
            }
        }

        const result: SlotField = {
            slotKey: key,
            label,
            value: isEmpty ? "---" : displayValue,
            rawValue: raw,
            isEmpty,
            lookupEntityType,
            lookupId,
            optionColor,
            dateText,
            timeText,
        };

        if (isSlotDebugEnabled()) {
            // Uniform per-slot trace covering every group (header/contact/address/details/grid/tags).
            // Keep payload small — full param dumps are noisy in mobile logs.
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const attrs = (param as any).attributes;
            console.log("[InfoCard.readSlot]", key, {
                type: param.type,
                rawType: typeof raw,
                isEmpty,
                value: result.value,
                lookupEntityType: result.lookupEntityType,
                lookupId: result.lookupId,
                hasFormatted: formatted != null,
                hasResolved: resolvedValue != null,
                attrs: attrs ? Object.keys(attrs) : null,
                attrLogicalName: attrs?.LogicalName ?? attrs?.logicalName ?? null,
                attrDisplayName: attrs?.DisplayName ?? attrs?.displayName ?? null,
            });
        }

        return result;
    }

    /**
     * Resolve a deterministic sample value for the form-designer preview.
     * Slot-name semantic match (titleField → "Adventure Works") wins over
     * type-based fallback so the preview reads like a real card rather than
     * a row of generic "Sample text" entries. Returns null only when neither
     * a slot override nor a type entry matches (caller treats that as "no sample").
     */
    private getDesignTimeSample(
        key: string,
        paramType: string,
    ): { value: string; raw: unknown } | null {
        const slotSample = SAMPLE_BY_SLOT[key];
        if (slotSample !== undefined) {
            return { value: slotSample, raw: slotSample };
        }
        const typeSample = SAMPLE_BY_TYPE[paramType];
        if (typeSample) {
            return typeSample;
        }
        return { value: "Sample", raw: "Sample" };
    }

    // ────────────────────────────────────────
    // Entity metadata (async label + option set resolution)
    // ────────────────────────────────────────

    /**
     * Fetch the current form record via WebAPI and match param values to columns.
     * This resolves labels and formatted values for bound fields where
     * param.attributes doesn't provide LogicalName (type-group properties).
     */
    // System/audit columns to exclude from value matching — these almost never
    // map to InfoCard slots and cause false positives (e.g., statuscode=1 vs bookingtype=1)
    private static readonly SYSTEM_COLUMNS = new Set([
        "statecode", "statuscode", "createdon", "modifiedon", "createdby", "modifiedby",
        "ownerid", "owningbusinessunit", "owningteam", "owninguser",
        "versionnumber", "timezoneruleversionnumber", "utcconversiontimezonecode",
        "importsequencenumber", "overriddencreatedon", "exchangerate",
    ]);

    /** Entities that store a color field on their record */
    private static readonly COLOR_FIELDS: Record<string, string> = {
        bookingstatus: "msdyn_statuscolor",
        msdyn_bookingstatus: "msdyn_statuscolor",
        msdyn_priority: "msdyn_prioritycolor",
    };

    /** Extract lookup entity type + ID from various PCF raw value formats */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private extractLookupRef(raw: any): { entityType?: string; id?: string } | null {
        if (Array.isArray(raw) && raw.length > 0 && raw[0].id) {
            return { entityType: raw[0].entityType ?? raw[0].etn, id: raw[0].id };
        }
        if (typeof raw === "object" && raw !== null && !Array.isArray(raw) && !(raw instanceof Date)) {
            const rawGuid = raw.Id?._formattedGuid ?? raw.Id?._rawGuid ?? raw.id;
            if (rawGuid) {
                return {
                    entityType: raw.LogicalName ?? raw.entityType,
                    id: this.formatGuid(String(rawGuid)),
                };
            }
        }
        return null;
    }

    /** Fetch a color field from a lookup entity record, normalizing hex format */
    private async fetchLookupColor(
        context: ComponentFramework.Context<IInputs>,
        entityType: string, id: string,
    ): Promise<string | undefined> {
        const colorField = InfoCard.COLOR_FIELDS[entityType];
        if (!colorField) return undefined;
        try {
            const rec = await context.webAPI.retrieveRecord(entityType, id, `?$select=${colorField}`);
            let color = rec[colorField];
            if (color && typeof color === "string" && !color.startsWith("#")) color = `#${color}`;
            if (color && typeof color === "string"
                && color.toLowerCase() !== "#ffffff" && color.toLowerCase() !== "#000000") {
                return color;
            }
        } catch { /* color fetch is optional */ }
        return undefined;
    }

    private async resolveRecordFieldsAsync(
        context: ComponentFramework.Context<IInputs>,
    ): Promise<Record<string, { label: string; value: string; color?: string }>> {
        if (!this._formEntityName || !this._formRecordId) return {};
        const overrides: Record<string, { label: string; value: string; color?: string }> = {};

        try {
            // Fetch full record — includes formatted value annotations for all columns
            const record = await context.webAPI.retrieveRecord(this._formEntityName, this._formRecordId);
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const rec = record as Record<string, any>;

            // Build column index: primitive columns only, excluding system fields
            const columnNames = Object.keys(rec).filter(k =>
                !k.includes("@") && !k.startsWith("_") && k !== "id"
                && !InfoCard.SYSTEM_COLUMNS.has(k)
            );

            // For each bound slot without a resolved column, match by value.
            // Plain numeric/boolean raws are too ambiguous to value-match (option-set integers,
            // two-options, durations all collide across columns) and the prior "first match wins"
            // behaviour produced silently-wrong column bindings. Skip them here — those slots rely
            // on attributes.LogicalName already populated by readSlot, or remain unresolved.
            // For non-numeric raws, only commit when exactly one column matches; multiple matches
            // are a sign of ambiguity and we refuse rather than guess.
            for (const key of ALL_SLOT_KEYS) {
                if (this._slotToColumn[key]) continue;
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const param = (context.parameters as Record<string, any>)[key];
                if (!param || param.type === "Unknown") continue;
                const raw = param.raw;
                if (raw === null || raw === undefined) continue;
                if (typeof raw === "string" && raw.startsWith("@")) continue;
                if (typeof raw === "number" || typeof raw === "boolean") continue;

                const candidates: string[] = [];
                for (const colName of columnNames) {
                    const colVal = rec[colName];
                    if (colVal === null || colVal === undefined) continue;
                    if (this.valuesMatch(raw, colVal)) candidates.push(colName);
                }
                if (candidates.length === 1) {
                    const colName = candidates[0];
                    this._slotToColumn[key] = colName;
                    const fmtKey = `${colName}@OData.Community.Display.V1.FormattedValue`;
                    const fmtVal = rec[fmtKey]
                        ?? (typeof rec.getFormattedValue === "function" ? rec.getFormattedValue(colName) : null);
                    if (fmtVal) {
                        this._resolvedValues[key] = String(fmtVal);
                    }
                } else if (candidates.length > 1) {
                    // Ambiguous match — common for date/datetime values on Bookings
                    // (starttime/scheduledstart/actualstart often share the same value
                    // by default). Tie-break ONLY by formatted-value comparison: if
                    // exactly one candidate's FormattedValue equals param.formatted,
                    // commit it; otherwise refuse rather than guess.
                    const paramFormatted = (param as Record<string, unknown>).formatted;
                    if (typeof paramFormatted === "string" && paramFormatted) {
                        const fmtMatches = candidates.filter(c => {
                            const fmt = rec[`${c}@OData.Community.Display.V1.FormattedValue`]
                                ?? (typeof rec.getFormattedValue === "function" ? rec.getFormattedValue(c) : null);
                            return fmt === paramFormatted;
                        });
                        if (fmtMatches.length === 1) {
                            const colName = fmtMatches[0];
                            this._slotToColumn[key] = colName;
                            this._resolvedValues[key] = paramFormatted;
                            continue;
                        }
                    }
                    // Last-resort tie-break (LABEL ONLY): when the maker bound a slot
                    // (param.type set, value is rendered directly from param.raw), the
                    // platform doesn't expose the column's LogicalName for usage="input"
                    // properties. We can't determine which column was bound. If this slot
                    // has a SLOT_PRESETS entry for the current entity AND that preset
                    // column is one of the candidates, commit it so the slot gets the
                    // correct DisplayName. The displayed VALUE is unaffected — it's
                    // already rendered from param.raw — only the label is disambiguated.
                    const presetCol: string | undefined = this._formEntityName
                        ? (SLOT_PRESETS[this._formEntityName] as Record<string, string> | undefined)?.[key]
                        : undefined;
                    if (presetCol && candidates.includes(presetCol)) {
                        this._slotToColumn[key] = presetCol;
                        const fmtKey = `${presetCol}@OData.Community.Display.V1.FormattedValue`;
                        const fmtVal = rec[fmtKey]
                            ?? (typeof rec.getFormattedValue === "function" ? rec.getFormattedValue(presetCol) : null);
                        if (fmtVal) this._resolvedValues[key] = String(fmtVal);
                        console.log(`[InfoCard] Ambiguous match for ${key} (${candidates.join(", ")}) — using preset hint "${presetCol}" for label only.`);
                        continue;
                    }
                    console.warn(`[InfoCard] Ambiguous value match for slot ${key} (matched ${candidates.join(", ")}); not committing — bind the slot to a column directly to disambiguate.`);
                }
            }

            console.log("[InfoCard] Record match:", Object.entries(this._slotToColumn)
                .map(([k, v]) => `${k}→${v}${this._resolvedValues[k] ? ` (${this._resolvedValues[k]})` : ""}`).join(", "));

            // Fetch metadata for all discovered columns (labels + option set defs)
            const allColumns = [...new Set(Object.values(this._slotToColumn))];
            if (allColumns.length > 0) {
                try {
                    const metadata = await context.utils.getEntityMetadata(this._formEntityName!, allColumns);
                    if (metadata?.Attributes) {
                        for (const attr of metadata.Attributes.getAll()) {
                            const logName = attr.LogicalName;
                            if (!logName) continue;
                            const entry: ColumnMeta = {
                                logicalName: logName,
                                displayName: this.extractDisplayNameFromMeta(attr.DisplayName) ?? this.formatLogicalName(logName),
                            };
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            const attrAny = attr as Record<string, any>;
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            const attrDesc = attrAny.attributeDescriptor as Record<string, any> | undefined;
                            // Capture Format ("duration", "phone", "email", etc.) so unbound
                            // preset slots — where param.type==="Unknown" — can still detect
                            // duration columns and format minutes as "Xd Xh Xm" instead of
                            // surfacing the raw integer. Different host versions expose this
                            // under different property names and as either string or numeric
                            // enum (1=Duration for IntegerAttributeMetadata).
                            let fmt: unknown = attrAny.Format ?? attrAny.FormatName
                                ?? attrDesc?.Format ?? attrDesc?.format
                                ?? attrDesc?.FormatName ?? attrDesc?.formatName;
                            if (typeof fmt === "number") {
                                // IntegerFormat enum: 0=None, 1=Duration, 2=DateAndTime, 3=TimeZone, 4=Language, 5=Locale
                                const intFormatMap: Record<number, string> = {
                                    1: "duration", 2: "datetime", 3: "timezone", 4: "language", 5: "locale",
                                };
                                fmt = intFormatMap[fmt];
                            }
                            if (typeof fmt === "string" && fmt.length > 0) {
                                entry.format = fmt.toLowerCase();
                            }
                            // Diagnostic: for duration columns (logName ends with "duration"
                            // or known Dataverse duration columns), dump the attr keys when
                            // we couldn't resolve a format. Helps identify the right key.
                            if (!entry.format && /duration$/i.test(logName) && isSlotDebugEnabled()) {
                                console.log(`[InfoCard.metadata] no Format for ${logName}; attr keys:`, Object.keys(attrAny), "attrDesc keys:", attrDesc ? Object.keys(attrDesc) : null);
                            }
                            const optionSet = attrDesc?.OptionSet ?? attrDesc?.optionSet;
                            if (optionSet?.Options) {
                                entry.options = [];
                                for (const opt of optionSet.Options) {
                                    const optLabel = opt.Label?.UserLocalizedLabel?.Label
                                        ?? opt.Label?.userLocalizedLabel?.label
                                        ?? opt.Label ?? String(opt.Value);
                                    const optColor = opt.Color && opt.Color !== "#0000ff" ? opt.Color : undefined;
                                    entry.options.push({ value: opt.Value, label: String(optLabel), color: optColor });
                                }
                            }
                            this._columnMetadata[logName] = entry;
                        }
                    }
                    console.log("[InfoCard] Bound field labels:", Object.entries(this._slotToColumn)
                        .map(([k, col]) => `${k}→"${this._columnMetadata[col]?.displayName ?? col}"`).join(", "));
                } catch { /* metadata is optional */ }
            }

            // Resolve formatted values from the fetched record for slots bound via
            // attributes.LogicalName (readSlot populated _slotToColumn but didn't pull
            // OData formatted values). Without this, attribute-bound numeric/option-set
            // slots produce empty override values even though the record carried the label.
            for (const [slotKey, colName] of Object.entries(this._slotToColumn)) {
                if (this._resolvedValues[slotKey]) continue;
                // Whole.Duration columns: Dataverse returns the raw integer in
                // minutes and the OData FormattedValue is unreliable across hosts
                // (some return "10,110" instead of "168h 30m"). Always re-format
                // from the raw integer when we recognize the slot as a duration.
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const param = (context.parameters as Record<string, any>)[slotKey];
                const paramType = String(param?.type ?? "");
                const isDuration = paramType === "Whole.Duration"
                    || String(param?.attributes?.Format ?? "").toLowerCase() === "duration"
                    || this._columnMetadata[colName]?.format === "duration"
                    || isLikelyDurationColumn(colName);
                if (isDuration) {
                    const rawMinutes = rec[colName];
                    if (typeof rawMinutes === "number") {
                        this._resolvedValues[slotKey] = this.formatDuration(rawMinutes, this._formatting, this._strings);
                        continue;
                    }
                }
                const fmtKey = `${colName}@OData.Community.Display.V1.FormattedValue`;
                const fmtVal = rec[fmtKey]
                    ?? (typeof rec.getFormattedValue === "function" ? rec.getFormattedValue(colName) : null);
                if (fmtVal) {
                    this._resolvedValues[slotKey] = String(fmtVal);
                }
            }

            // Resolve colors for lookup-based tag fields
            for (const key of TAG_KEYS) {
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const param = (context.parameters as Record<string, any>)[key];
                if (!param) continue;
                const raw = param.raw;
                if (!raw || (typeof raw === "string" && raw.startsWith("@"))) continue;
                const ref = this.extractLookupRef(raw);
                if (ref?.entityType && ref?.id) {
                    const color = await this.fetchLookupColor(context, ref.entityType, ref.id);
                    if (color) this._resolvedColors[key] = color;
                }
            }

            // Build overrides map for React to apply
            // Skip titleField — it's usage="bound" and already has correct formatted value
            for (const [slotKey, colName] of Object.entries(this._slotToColumn)) {
                if (slotKey === "titleField") continue;
                const meta = this._columnMetadata[colName];
                const fmtValue = this._resolvedValues[slotKey];
                const color = this._resolvedColors[slotKey];
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                const param = (context.parameters as Record<string, any>)[slotKey];
                const attrs = param?.attributes;
                const hasMakerBinding = !!(attrs?.LogicalName ?? attrs?.logicalName);
                // Surface the real column DisplayName when (a) the maker bound
                // the slot directly, (b) it came from a SLOT_PRESETS entry, OR
                // (c) entity metadata gave us a real DisplayName for the resolved
                // column. Case (c) covers value-matched slots — usage="input"
                // properties (grids, tags, details, subtitles) get attributes={}
                // from the platform regardless of binding, so the value-matcher
                // is the only way to learn the column. Without surfacing the
                // metadata DisplayName here, every grid slot renders as
                // "Grid 1/2/3" instead of "Start Time/End Time/Duration".
                const hasMetadataDisplayName = !!meta?.displayName
                    && meta.displayName !== this.formatLogicalName(colName);
                const useRealLabel = hasMakerBinding
                    || this._presetSlots.has(slotKey)
                    || hasMetadataDisplayName;
                if (meta || fmtValue || color) {
                    // Only emit a label when we have a "real" one. The generic fallback
                    // (formatPropertyKey -> "Grid 1") is no improvement over the UI's
                    // own default and — because override merge is shallow per-slot — a
                    // sparse late pass emitting "Grid 1" would clobber a real label
                    // ("Start Time") set by an earlier richer pass. When useRealLabel
                    // is false, omit label and let the React applyOverride keep the
                    // existing slot label (`ov.label || f.label`).
                    overrides[slotKey] = {
                        label: useRealLabel
                            ? (meta?.displayName ?? this.formatLogicalName(colName))
                            : "",
                        value: fmtValue ?? "",
                        color,
                    };
                }
            }

            // Include tag colors even if the tag wasn't matched by value
            // (lookups like bookingstatus can't be matched via _column_value filtering)
            for (const key of TAG_KEYS) {
                const color = this._resolvedColors[key];
                if (color && !overrides[key]) {
                    overrides[key] = { label: "", value: "", color };
                } else if (color && overrides[key]) {
                    overrides[key].color = color;
                }
            }

            console.log("[InfoCard] Record overrides:", Object.entries(overrides)
                .map(([k, v]) => `${k}→"${v.label}": ${v.value}${v.color ? ` [${v.color}]` : ""}`).join(", "));
            return overrides;
        } catch (err) {
            console.warn("[InfoCard] resolveRecordFields failed:", err);
            return overrides;
        }
    }

    /** Match a PCF param raw value against a WebAPI record column value */
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private valuesMatch(paramRaw: any, recordVal: any): boolean {
        // Number (OptionSet, integer, decimal)
        if (typeof paramRaw === "number" && typeof recordVal === "number") {
            return paramRaw === recordVal;
        }
        // Date: param is Date object, record is ISO string
        if (paramRaw instanceof Date && typeof recordVal === "string") {
            const recordDate = new Date(recordVal);
            return !isNaN(recordDate.getTime()) && Math.abs(paramRaw.getTime() - recordDate.getTime()) < 60000;
        }
        // String
        if (typeof paramRaw === "string" && typeof recordVal === "string") {
            return paramRaw === recordVal;
        }
        // Lookup array: compare ID
        if (Array.isArray(paramRaw) && paramRaw.length > 0 && paramRaw[0].id && typeof recordVal === "string") {
            return paramRaw[0].id.toLowerCase() === recordVal.toLowerCase();
        }
        return false;
    }

    private extractDisplayNameFromMeta(dn: unknown): string | null {
        if (typeof dn === "string" && dn.length > 0) return dn;
        if (dn && typeof dn === "object") {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const obj = dn as Record<string, any>;
            return obj.UserLocalizedLabel?.Label
                ?? obj.userLocalizedLabel?.label
                ?? obj.LocalizedLabels?.[0]?.Label
                ?? obj.localizedLabels?.[0]?.label
                ?? null;
        }
        return null;
    }

    // ────────────────────────────────────────
    // Related fields
    // ────────────────────────────────────────

    /**
     * Detect related field references using the @ convention.
     *
     * Two source types:
     *   @fieldName       → fetch from title field's lookup entity (work order)
     *   @nav.field       → fetch nav.field from title entity via $expand
     *   @.nav.field      → fetch from CURRENT form record via $expand
     *
     * The @. prefix navigates from the current record instead of the title entity.
     */
    private detectRelatedFields(context: ComponentFramework.Context<IInputs>): RelatedFieldMapping[] {
        const mappings: RelatedFieldMapping[] = [];
        const scanKeys = ALL_SLOT_KEYS.filter(k => k !== "titleField");

        for (const slotKey of scanKeys) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const param = (context.parameters as Record<string, any>)[slotKey];
            if (!param) continue;
            const raw = param.raw;
            if (typeof raw !== "string" || !raw.startsWith("@")) continue;

            let fetchField: string;
            let sourceSlot: string;

            if (raw.startsWith("@.")) {
                // @. prefix → navigate from current form record
                fetchField = raw.substring(2).trim();
                sourceSlot = "__currentRecord__";
            } else {
                // @ prefix → navigate from title entity
                fetchField = raw.substring(1).trim();
                sourceSlot = "titleField";
            }

            if (fetchField) {
                mappings.push({ sourceSlot, fetchField, targetSlot: slotKey });
            }
        }

        return mappings;
    }

    /**
     * Read a field value from a WebAPI record, handling lookups, direct fields, and images.
     * Tries OData annotation properties and getFormattedValue() method.
     */
    private readField(record: ComponentFramework.WebApi.Entity, col: string): {
        value: string; label: string; lookupId?: string; lookupEntityType?: string;
    } | null {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const rec = record as Record<string, any>;
        const getFormatted = typeof rec.getFormattedValue === "function"
            ? (key: string) => rec.getFormattedValue(key) as string | null
            : () => null;
        // Readable seed for `label` — gets overridden by entity metadata in step 6a
        // when available, but stays human-friendly ("Work Order Instructions") if
        // the metadata fetch is skipped or fails silently.
        const seedLabel = this.formatLogicalName(col);

        // Image columns
        if (col === "entityimage" || col.endsWith("image") || col.endsWith("_image")) {
            const allImageKeys = Object.keys(rec).filter(k =>
                k.toLowerCase().includes("image") && !k.includes("@") && !k.includes("timestamp"));
            for (const key of allImageKeys) {
                if (key.toLowerCase().endsWith("_url") && rec[key]) {
                    return { value: String(rec[key]), label: "image" };
                }
            }
            for (const key of allImageKeys) {
                const val = rec[key];
                if (val && typeof val === "string" && (val as string).length > 100 && !key.endsWith("_url")) {
                    return { value: `data:image/png;base64,${val}`, label: "image" };
                }
            }
            return null;
        }

        // Lookup pattern: _col_value
        const lookupKey = `_${col}_value`;
        const lookupFormatted = rec[`${lookupKey}@OData.Community.Display.V1.FormattedValue`]
            ?? rec[`${col}@OData.Community.Display.V1.FormattedValue`]
            ?? getFormatted(lookupKey)
            ?? getFormatted(col);
        const lookupVal = rec[lookupKey];
        if (lookupFormatted) {
            const etnKey = `${lookupKey}@Microsoft.Dynamics.CRM.lookuplogicalname`;
            return {
                value: String(lookupFormatted),
                label: seedLabel,
                lookupId: lookupVal ? String(lookupVal) : undefined,
                lookupEntityType: rec[etnKey] ? String(rec[etnKey]) : undefined,
            };
        }
        if (lookupVal != null) {
            return { value: String(lookupVal), label: seedLabel, lookupId: String(lookupVal) };
        }

        // Direct field with formatted value
        const formatted = rec[`${col}@OData.Community.Display.V1.FormattedValue`]
            ?? getFormatted(col);
        if (formatted) return { value: String(formatted), label: seedLabel };

        const direct = rec[col];
        if (direct == null) return null;

        // Reject non-primitive direct values — happens when OData $select returns an
        // expanded navigation object (e.g. `$select=msdyn_serviceaccount` yielding the
        // related entity) rather than a flat primitive. `String({...})` would render
        // "[object Object]". Try to recover a display name from common fields, else
        // skip so the lookup-primitive path can take over.
        if (typeof direct === "object") {
            const obj = direct as Record<string, unknown>;
            const candidate = obj.name ?? obj.fullname ?? obj.title ?? obj.subject;
            if (typeof candidate === "string" && candidate.length > 0) {
                return { value: candidate, label: seedLabel };
            }
            // Fallback: scan for any `*name` / `*Name` string property
            // (e.g. accountname, primarycontactname). Skip annotation keys
            // and odata metadata.
            for (const key of Object.keys(obj)) {
                if (key.includes("@") || key.startsWith("_")) continue;
                if (!/name$/i.test(key)) continue;
                const v = obj[key];
                if (typeof v === "string" && v.length > 0) {
                    return { value: v, label: seedLabel };
                }
            }
            return null;
        }
        return { value: String(direct), label: seedLabel };
    }

    /**
     * Fetch related fields using OData with layered fallback.
     *
     * For title-entity fetches (@syntax): tries $expand for dotted paths, falls back to hop2.
     * For current-record fetches (@. syntax): skips $expand entirely (column logical names
     * aren't OData navigation properties) and goes straight to hop2.
     */
    private async fetchRelatedFields(
        context: ComponentFramework.Context<IInputs>,
        entityType: string,
        id: string,
        columns: string[],
        skipExpand = false,
    ): Promise<Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string; color?: string }>> {
        try {
            console.log("[InfoCard] fetchRelatedFields:", entityType, id, columns, skipExpand ? "(hop2 only)" : "");
            const results: Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string; color?: string }> = {};

            // ── 1. Classify columns ──────────────────────────────────────
            const directCols: string[] = [];
            const expandGroups: Record<string, string[]> = {};
            // For @. (current-record), dotted paths use hop2 not $expand
            const hop2Groups: Record<string, string[]> = {};

            for (const col of columns) {
                const dotIdx = col.indexOf(".");
                if (dotIdx > 0) {
                    const navProp = col.substring(0, dotIdx);
                    const field = col.substring(dotIdx + 1);
                    if (skipExpand) {
                        // Go straight to hop2 — column names aren't OData nav properties
                        if (!hop2Groups[navProp]) hop2Groups[navProp] = [];
                        if (!hop2Groups[navProp].includes(field)) hop2Groups[navProp].push(field);
                        // Don't add to directCols — hop2 fetches _navProp_value itself
                    } else {
                        if (!expandGroups[navProp]) expandGroups[navProp] = [];
                        if (!expandGroups[navProp].includes(field)) expandGroups[navProp].push(field);
                        if (!directCols.includes(navProp)) directCols.push(navProp);
                    }
                } else {
                    if (!directCols.includes(col)) directCols.push(col);
                }
            }

            const hasExpand = Object.keys(expandGroups).length > 0;

            // ── 2. Build combined query ──────────────────────────────────
            const selectParts = directCols;
            const expandParts = Object.entries(expandGroups).map(
                ([nav, fields]) => `${nav}($select=${fields.join(",")})`
            );
            let options = `?$select=${selectParts.join(",")}`;
            if (hasExpand) {
                options += `&$expand=${expandParts.join(",")}`;
            }

            // ── 3. Execute with layered fallback ─────────────────────────
            let record: ComponentFramework.WebApi.Entity = {};
            let usedFallback = false;

            try {
                record = await context.webAPI.retrieveRecord(entityType, id, options);
            } catch (expandErr) {
                console.warn("[InfoCard] Combined query failed — using fallback:", expandErr);
                usedFallback = true;
                record = {};

                // Fallback A: fetch direct columns per-column
                for (const col of directCols) {
                    // Try both primitive and lookup patterns
                    for (const selectCol of [col, `_${col}_value`]) {
                        try {
                            const r = await context.webAPI.retrieveRecord(entityType, id, `?$select=${selectCol}`);
                            for (const k of Object.keys(r)) { record[k] = r[k]; }
                        } catch (e) {
                            // Expected for non-lookup `_xxx_value` and unknown columns; debug-only.
                            if (isSlotDebugEnabled()) console.log(`[InfoCard] per-column fetch ${selectCol} failed:`, (e as Error)?.message);
                        }
                    }
                }

                // Fallback B: try each $expand group individually
                for (const [nav, fields] of Object.entries(expandGroups)) {
                    try {
                        const r = await context.webAPI.retrieveRecord(
                            entityType, id, `?$expand=${nav}($select=${fields.join(",")})`
                        );
                        if (r[nav]) record[nav] = r[nav];
                    } catch {
                        // Fallback C: legacy hop2 — resolve lookup, then fetch from related entity
                        console.warn(`[InfoCard] $expand(${nav}) failed — trying hop2`);
                        const hop1 = this.readField(record, nav);
                        if (hop1?.lookupId && hop1?.lookupEntityType) {
                            try {
                                const hop2Record = await context.webAPI.retrieveRecord(
                                    hop1.lookupEntityType, hop1.lookupId, `?$select=${fields.join(",")}`
                                );
                                record[nav] = hop2Record;
                            } catch {
                                console.warn(`[InfoCard] Hop2 for ${nav} failed — skipping`);
                            }
                        }
                    }
                }
            }

            // ── 3b. Follow-up: ensure lookup primitives are fetched ──────
            // Chase the `_col_value` form for columns that came back with NO
            // entry in the response, OR that came back as a sparse expanded
            // navigation object (Dataverse sometimes returns `{ "@odata.id": ... }`
            // for `?$select=navProp` instead of the lookup primitive). Without
            // this, `@msdyn_serviceaccount` resolves to an object readField can't
            // recover a display name from, and the slot renders empty.
            //
            // Guard: if `col in record` AS A PRIMITIVE (string/number/etc) or
            // `_col_value` is already present, skip — fetching `_xxx_value` for
            // a non-lookup column produces an OData 400.
            if (!usedFallback) {
                const missingLookupCols = directCols.filter(col => {
                    const lookupKeyPresent = Object.prototype.hasOwnProperty.call(record, `_${col}_value`);
                    if (lookupKeyPresent) return false;
                    const keyPresent = Object.prototype.hasOwnProperty.call(record, col);
                    if (!keyPresent) return true;
                    // Key present but value is a non-null object (sparse expand) —
                    // we still need the lookup primitive for FormattedValue.
                    const v = record[col];
                    return v !== null && typeof v === "object";
                });

                if (missingLookupCols.length > 0) {
                    const lookupSelect = missingLookupCols.map(col => `_${col}_value`).join(",");
                    try {
                        const lookupRecord = await context.webAPI.retrieveRecord(entityType, id, `?$select=${lookupSelect}`);
                        for (const k of Object.keys(lookupRecord)) { record[k] = lookupRecord[k]; }
                    } catch {
                        for (const col of missingLookupCols) {
                            try {
                                const r = await context.webAPI.retrieveRecord(entityType, id, `?$select=_${col}_value`);
                                for (const k of Object.keys(r)) { record[k] = r[k]; }
                            } catch { /* skip */ }
                        }
                    }
                }
            }

            // ── 4. Parse direct fields ───────────────────────────────────
            for (const col of directCols) {
                const isExpandIntermediary = expandGroups[col] && !columns.includes(col);
                const result = this.readField(record, col);
                if (result && !isExpandIntermediary) {
                    results[col] = result;
                }
            }

            // ── 5. Parse expanded (nested) fields ────────────────────────
            for (const [nav, fields] of Object.entries(expandGroups)) {
                const nested = record[nav];
                const nestedObj = (nested && typeof nested === "object" && !Array.isArray(nested))
                    ? nested as Record<string, unknown> : null;

                if (nestedObj) {
                    for (const field of fields) {
                        const result = this.readField(nestedObj as ComponentFramework.WebApi.Entity, field);
                        if (result) {
                            results[`${nav}.${field}`] = result;
                        }
                    }
                }

                // Hop2 fallback: if expanded parsing missed fields
                const missingFields = fields.filter(f => !results[`${nav}.${f}`]);
                if (missingFields.length > 0) {
                    const hop1 = this.readField(record, nav);
                    if (hop1?.lookupId && hop1?.lookupEntityType) {
                        try {
                            const hop2Record = await context.webAPI.retrieveRecord(
                                hop1.lookupEntityType, hop1.lookupId, `?$select=${missingFields.join(",")}`
                            );
                            for (const field of missingFields) {
                                const result = this.readField(hop2Record, field);
                                if (result) {
                                    results[`${nav}.${field}`] = result;
                                }
                            }
                        } catch { /* hop2 failed */ }
                    } else {
                        // No lookup primitive yet — try fetching it
                        try {
                            const lookupRecord = await context.webAPI.retrieveRecord(entityType, id, `?$select=_${nav}_value`);
                            const hop1b = this.readField(lookupRecord, nav);
                            if (hop1b?.lookupId && hop1b?.lookupEntityType) {
                                const hop2Record = await context.webAPI.retrieveRecord(
                                    hop1b.lookupEntityType, hop1b.lookupId, `?$select=${missingFields.join(",")}`
                                );
                                for (const field of missingFields) {
                                    const result = this.readField(hop2Record, field);
                                    if (result) {
                                        results[`${nav}.${field}`] = result;
                                    }
                                }
                            }
                        } catch { /* hop2b failed */ }
                    }
                }
            }

            // ── 5b. Direct hop2 for @. dotted paths (skip $expand) ────────
            for (const [nav, fields] of Object.entries(hop2Groups)) {
                // Resolve the lookup column to get target entity + ID
                const hop1 = this.readField(record, nav);
                if (hop1?.lookupId && hop1?.lookupEntityType) {
                    try {
                        const hop2Record = await context.webAPI.retrieveRecord(
                            hop1.lookupEntityType, hop1.lookupId, `?$select=${fields.join(",")}`
                        );
                        for (const field of fields) {
                            const result = this.readField(hop2Record, field);
                            if (result) {
                                results[`${nav}.${field}`] = result;
                            }
                        }
                    } catch { /* hop2 failed */ }
                } else {
                    // Lookup primitive not in record yet — fetch it
                    try {
                        const lookupRecord = await context.webAPI.retrieveRecord(entityType, id, `?$select=_${nav}_value`);
                        // Merge into record so step 6b can resolve the target entity
                        for (const k of Object.keys(lookupRecord)) { record[k] = lookupRecord[k]; }
                        const hop1b = this.readField(lookupRecord, nav);
                        if (hop1b?.lookupId && hop1b?.lookupEntityType) {
                            const hop2Record = await context.webAPI.retrieveRecord(
                                hop1b.lookupEntityType, hop1b.lookupId, `?$select=${fields.join(",")}`
                            );
                            for (const field of fields) {
                                const result = this.readField(hop2Record, field);
                                if (result) {
                                    results[`${nav}.${field}`] = result;
                                }
                            }
                        }
                    } catch { /* hop2b failed */ }
                }
            }

            // ── 6. Resolve display names from entity metadata ────────────
            // 6a. Direct fields on the source entity
            try {
                const colsToResolve = Object.keys(results).filter(k => !k.startsWith("__") && !k.includes("."));
                if (colsToResolve.length > 0) {
                    const entityMeta = await context.utils.getEntityMetadata(entityType, colsToResolve);
                    if (entityMeta?.Attributes) {
                        for (const attr of entityMeta.Attributes.getAll()) {
                            const logicalName = attr.LogicalName;
                            if (!logicalName || !results[logicalName]) continue;
                            const dn = this.extractDisplayNameFromMeta(attr.DisplayName);
                            if (dn) results[logicalName].label = dn;
                            // Detect Whole.Duration columns and re-format the value from the
                            // raw integer in `record`. The OData FormattedValue is unreliable
                            // for duration columns (often returns the locale-grouped integer
                            // "10,110" instead of "168h 30m"), so we override with our own
                            // formatter when metadata confirms the column is a duration.
                            // eslint-disable-next-line @typescript-eslint/no-explicit-any
                            const attrAny = attr as Record<string, any>;
                            const attrDesc = attrAny.attributeDescriptor ?? attrAny.AttributeDescriptor ?? attrAny;
                            const fmt = String(attrDesc?.Format ?? attrDesc?.format ?? "").toLowerCase();
                            if (fmt === "duration") {
                                const rawMin = (record as Record<string, unknown>)[logicalName];
                                if (typeof rawMin === "number" && Number.isFinite(rawMin)) {
                                    results[logicalName].value = this.formatDuration(rawMin, this._formatting, this._strings);
                                }
                            }
                        }
                    }
                }
            } catch { /* metadata fetch is optional */ }

            // 6a-fallback: Heuristic duration detection when metadata didn't fire (or
            // didn't expose Format). Catches `duration`, `msdyn_duration`, etc. so the
            // fetched value renders as "Xh Ym" instead of "10,110".
            for (const col of Object.keys(results)) {
                if (col.startsWith("__") || col.includes(".")) continue;
                if (!isUnbindableColumn(col)) continue;
                if (!/(^|_)duration(minutes)?$/i.test(col)) continue;
                const current = results[col];
                if (!current) continue;
                // Already looks like a formatted duration ("Xh Ym" / "X hours") — leave alone.
                if (/[a-z]/i.test(current.value) && !/^[\d,.\s-]+$/.test(current.value)) continue;
                const rawMin = (record as Record<string, unknown>)[col];
                let minutes: number | null = null;
                if (typeof rawMin === "number" && Number.isFinite(rawMin)) {
                    minutes = rawMin;
                } else {
                    const parsed = parseInt(String(current.value).replace(/[^\d-]/g, ""), 10);
                    if (Number.isFinite(parsed)) minutes = parsed;
                }
                if (minutes !== null) {
                    current.value = this.formatDuration(minutes, this._formatting, this._strings);
                }
            }

            // 6b. Dotted-path fields: resolve from the target entity
            const allDottedGroups = { ...expandGroups, ...hop2Groups };
            for (const [nav, fields] of Object.entries(allDottedGroups)) {
                const parsedFields = fields.filter(f => results[`${nav}.${f}`]);
                if (parsedFields.length === 0) continue;
                // Determine target entity from the nav property's lookup metadata
                const hop1 = this.readField(record, nav);
                const targetEntity = hop1?.lookupEntityType;
                if (!targetEntity) continue;
                try {
                    const navMeta = await context.utils.getEntityMetadata(targetEntity, parsedFields);
                    if (navMeta?.Attributes) {
                        for (const attr of navMeta.Attributes.getAll()) {
                            const logicalName = attr.LogicalName;
                            const dn = this.extractDisplayNameFromMeta(attr.DisplayName);
                            if (logicalName && dn && results[`${nav}.${logicalName}`]) {
                                results[`${nav}.${logicalName}`].label = dn;
                            }
                        }
                    }
                } catch { /* metadata fetch is optional */ }
            }

            // ── 7. Resolve colors for lookup results from known color entities ──
            for (const key of Object.keys(results)) {
                const val = results[key];
                if (!val.lookupId || !val.lookupEntityType) continue;
                const color = await this.fetchLookupColor(context, val.lookupEntityType, val.lookupId);
                if (color) results[key] = { ...val, color };
            }

            const resultSummary = Object.entries(results).map(([k, v]) => `${k}="${v.value}" (${v.label})`).join(", ");
            console.log(`[InfoCard] fetchRelatedFields done: ${resultSummary}`);

            return results;
        } catch (err) {
            console.error("[InfoCard] fetchRelatedFields ERROR:", err);
            return {};
        }
    }

    // ────────────────────────────────────────
    // Binding diagnostics (design-time panel)
    // ────────────────────────────────────────

    private buildBindingDiagnostics(
        context: ComponentFramework.Context<IInputs>,
        mappings: RelatedFieldMapping[],
    ): BindingDiagnostic[] {
        const diagnostics: BindingDiagnostic[] = [];

        for (const key of ALL_SLOT_KEYS) {
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            const param = (context.parameters as Record<string, any>)[key];
            if (!param) continue;

            const paramType = param.type;
            const raw = param.raw;

            // Skip truly unconfigured slots
            if (!paramType && (raw === null || raw === undefined)) continue;
            if (paramType === "Unknown") continue;

            const entry: BindingDiagnostic = {
                slotKey: key,
                slotLabel: SLOT_LABELS[key] ?? key,
                bindingType: "bound",
                rawExpression: "",
            };

            if (typeof raw === "string" && raw.startsWith("@.")) {
                entry.bindingType = "current-related";
                entry.rawExpression = raw;
                const path = raw.substring(2).trim();
                if (!path) {
                    entry.warning = "Empty path after @. — expected @.navProp.field";
                } else if (!path.includes(".") && !isUnbindableColumn(path)) {
                    entry.warning = `Direct field '${path}' on current record — consider $-binding (faster, design-time validation). Keep @.${path} if the column type isn't bindable.`;
                }
            } else if (typeof raw === "string" && raw.startsWith("@")) {
                entry.bindingType = "title-related";
                entry.rawExpression = raw;
                const path = raw.substring(1).trim();
                if (!path) {
                    entry.warning = "Empty path after @ — expected @fieldName";
                }
            } else {
                entry.bindingType = "bound";
                const logicalName = param.attributes?.LogicalName ?? param.attributes?.logicalName;
                entry.rawExpression = logicalName ?? `[${paramType ?? "unknown type"}]`;
            }

            // Only include entries with warnings (syntax issues)
            if (entry.warning) {
                diagnostics.push(entry);
            }
        }

        return diagnostics;
    }

    // ────────────────────────────────────────
    // Duration formatting
    // ────────────────────────────────────────

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    private isDurationField(attrs: Record<string, any>, raw: unknown, formatted: string | undefined): boolean {
        if (typeof raw !== "number") return false;
        const fmt = String(attrs?.Format ?? "").toLowerCase();
        if (fmt === "duration") return true;
        if (formatted) {
            const f = formatted.toLowerCase().trim();
            if (f.endsWith("hours") || f.endsWith("hour") || f.endsWith("minutes") || f.endsWith("minute")) {
                return true;
            }
        }
        return false;
    }

    private formatDuration(
        minutes: number,
        formatting?: MaybeFormatting,
        strings?: InfoCardStrings,
    ): string {
        return formatLocalizedDuration(
            minutes,
            strings ?? DEFAULT_STRINGS,
            formatting?.formatInteger,
        );
    }

    // ────────────────────────────────────────
    // Helpers
    // ────────────────────────────────────────

    /** Format a compact GUID (no hyphens) into standard 8-4-4-4-12 format */
    private formatGuid(raw: string): string {
        const hex = raw.replace(/[{}-]/g, "");
        if (hex.length !== 32) return raw;
        return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
    }

    /** Convert property key like "gridField1" to "Grid 1" */
    private formatPropertyKey(key: string): string {
        return key
            .replace(/Field(\d*)$/, " $1")
            .replace(/([a-z])([A-Z])/g, "$1 $2")
            .replace(/^\w/, c => c.toUpperCase())
            .trim();
    }

    /** Convert Dataverse logical name to readable label: "msdyn_workordertype" → "Work Order Type" */
    private formatLogicalName(logicalName: string): string {
        return logicalName
            .replace(/^(msdyn_|new_|cr[a-z0-9]+_|cli_|smp_)/, "")
            .replace(/_/g, " ")
            .replace(/([a-z])([A-Z])/g, "$1 $2")
            .replace(/\b\w/g, c => c.toUpperCase())
            .trim();
    }

    // ────────────────────────────────────────
    // Config resolution
    // ────────────────────────────────────────

    private resolveLayout(context: ComponentFramework.Context<IInputs>): LayoutMode {
        const raw = context.parameters.layout?.raw;
        if (raw === "smart" || raw === "contact" || raw === "compact") return raw;
        return "contact"; // default
    }

    private resolveTheme(context: ComponentFramework.Context<IInputs>): InfoCardTheme {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        const fdl = (context as Record<string, any>).fluentDesignLanguage;
        if (fdl?.tokenTheme) {
            const t = fdl.tokenTheme;
            const dt = defaultTheme.typography;
            // Fluent v9 ramp tokens — sizes/lineHeights are CSS strings ("14px"); weights are numbers.
            const sizeBase100 = t.fontSizeBase100 ?? dt.versionBadge.fontSize;
            const sizeBase200 = t.fontSizeBase200 ?? dt.fieldLabel.fontSize;
            const sizeBase300 = t.fontSizeBase300 ?? dt.body.fontSize;
            const sizeBase400 = t.fontSizeBase400 ?? dt.title.fontSize;
            const lhBase100 = t.lineHeightBase100 ?? dt.versionBadge.lineHeight;
            const lhBase200 = t.lineHeightBase200 ?? dt.fieldLabel.lineHeight;
            const lhBase300 = t.lineHeightBase300 ?? dt.body.lineHeight;
            const lhBase400 = t.lineHeightBase400 ?? dt.title.lineHeight;
            const wRegular = Number(t.fontWeightRegular ?? dt.body.fontWeight);
            const wMedium = Number(t.fontWeightMedium ?? dt.tag.fontWeight);
            const wSemibold = Number(t.fontWeightSemibold ?? dt.title.fontWeight);
            return {
                cardBg: t.colorNeutralBackground1 ?? defaultTheme.cardBg,
                textPrimary: t.colorNeutralForeground1 ?? defaultTheme.textPrimary,
                textSecondary: t.colorNeutralForeground3 ?? defaultTheme.textSecondary,
                textMuted: t.colorNeutralForeground4 ?? defaultTheme.textMuted,
                border: t.colorNeutralStroke1 ?? defaultTheme.border,
                borderLight: t.colorNeutralStroke2 ?? defaultTheme.borderLight,
                brand: t.colorBrandForeground1 ?? defaultTheme.brand,
                brandLight: t.colorBrandBackground2 ?? defaultTheme.brandLight,
                radius: t.borderRadiusMedium ?? defaultTheme.radius,
                shadow: t.shadow4 ?? defaultTheme.shadow,
                fontFamily: t.fontFamilyBase ?? defaultTheme.fontFamily,
                typography: {
                    title:           { fontSize: sizeBase400, fontWeight: wSemibold, lineHeight: lhBase400 },
                    subtitle:        { fontSize: sizeBase300, fontWeight: wRegular,  lineHeight: lhBase300 },
                    body:            { fontSize: sizeBase300, fontWeight: wRegular,  lineHeight: lhBase300 },
                    groupLabel:      { fontSize: sizeBase200, fontWeight: wSemibold, lineHeight: lhBase200, textTransform: "uppercase", letterSpacing: "0.04em" },
                    fieldLabel:      { fontSize: sizeBase200, fontWeight: wRegular,  lineHeight: lhBase200 },
                    inlineBoldLabel: { fontSize: sizeBase300, fontWeight: wSemibold, lineHeight: lhBase300 },
                    tag:             { fontSize: sizeBase200, fontWeight: wMedium,   lineHeight: lhBase200 },
                    avatarInitials:  { fontSize: sizeBase300, fontWeight: wSemibold, lineHeight: lhBase300 },
                    versionBadge:    { fontSize: sizeBase100, fontWeight: wRegular,  lineHeight: lhBase100 },
                },
            };
        }
        return defaultTheme;
    }
}
