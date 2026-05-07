/**
 * InfoCard — A compact info card React component for PCF virtual controls.
 *
 * Renders slot-based fields in three layout modes: smart (collapsible),
 * contact (full card), and compact (dense grid). All styling is inline,
 * all icons are inline SVGs. No external dependencies.
 */
import * as React from "react";

// ════════════════════════════════════════════════════════════════════
// URL safety
// ════════════════════════════════════════════════════════════════════

/**
 * Validate a user-provided URL string before using it in an href attribute.
 * React 16.8.6 does NOT block javascript:/data:/vbscript: URLs (the protections
 * landed in 16.9 / 17). We must do it ourselves.
 *
 * @returns the safe URL to use in href, or null if the input must be rendered as plain text.
 */
export function safeHttpUrl(value: string | undefined | null): string | null {
    if (!value) return null;
    const trimmed = String(value).trim();
    if (!trimmed) return null;
    // Strip ASCII control chars that some URL parsers ignore (\t, \n, \r in href)
    // eslint-disable-next-line no-control-regex
    const cleaned = trimmed.replace(/[\u0000-\u001F\u007F]/g, "");
    if (!cleaned) return null;
    // If the value already declares a scheme, only allow http(s).
    const schemeMatch = /^([a-zA-Z][a-zA-Z0-9+.-]*):/.exec(cleaned);
    if (schemeMatch) {
        const scheme = schemeMatch[1].toLowerCase();
        if (scheme === "http" || scheme === "https") return cleaned;
        return null; // javascript:, data:, vbscript:, file:, ftp:, etc.
    }
    // Reject protocol-relative URLs ("//evil.com/...") — ambiguous, often unsafe.
    if (cleaned.startsWith("//")) return null;
    // No scheme — assume https.
    return `https://${cleaned}`;
}

// ════════════════════════════════════════════════════════════════════
// Types
// ════════════════════════════════════════════════════════════════════

export type LayoutMode = "smart" | "contact" | "compact";

export interface SlotField {
    /** Slot identifier (e.g. "gridField2", "tagField1") — set by readSlot or mergeRelatedFields.
     *  Used by the override pass to apply record-fetched labels/colors back onto the right
     *  rendered row, regardless of how readSlotGroup compacts unconfigured slots. */
    slotKey?: string;
    label: string;
    value: string;
    rawValue: unknown;
    isEmpty: boolean;
    /** True while related-record fetch is in flight; renders as a shimmer placeholder. */
    isPending?: boolean;
    /** True for slots populated by SLOT_PRESETS (auto-fill for known entity types).
     *  Preset slots are hidden when empty even with hideEmpty=false — they're speculative
     *  fills for columns that may not have data on this record. */
    isPreset?: boolean;
    lookupEntityType?: string;
    lookupId?: string;
    /** OptionSet color from Dataverse metadata (hex, e.g. "#d13438") */
    optionColor?: string;
    /** Pre-computed date portion (locale-formatted) for DateAndTime.DateAndTime fields. */
    dateText?: string;
    /** Pre-computed time portion (locale-formatted) for DateAndTime.DateAndTime fields.
     *  When both dateText and timeText are present, the renderer stacks them on two lines. */
    timeText?: string;
}

export interface InfoCardData {
    title: SlotField | null;
    subtitles: SlotField[];
    phones: SlotField[];
    email: SlotField | null;
    web: SlotField | null;
    address: SlotField | null;
    latitude: number | null;
    longitude: number | null;
    details: SlotField[];
    gridFields: SlotField[];
    tags: SlotField[];
    imageUrl: string | null;
}

export interface InfoCardTheme {
    cardBg: string;
    textPrimary: string;
    textSecondary: string;
    textMuted: string;
    border: string;
    borderLight: string;
    brand: string;
    brandLight: string;
    radius: string;
    shadow: string;
    fontFamily: string;
}

export const defaultTheme: InfoCardTheme = {
    cardBg: "#ffffff",
    textPrimary: "#242424",
    textSecondary: "#616161",
    textMuted: "#8a8a8a",
    border: "#e0e0e0",
    borderLight: "#f0f0f0",
    brand: "#0f6cbd",
    brandLight: "#e8f1fa",
    radius: "8px",
    shadow: "0 1px 3px rgba(0,0,0,0.08), 0 0 1px rgba(0,0,0,0.04)",
    fontFamily: "'Segoe UI', -apple-system, BlinkMacSystemFont, sans-serif",
};

export interface RelatedFieldMapping {
    sourceSlot: string;
    fetchField: string;
    targetSlot: string;
}

export interface BindingDiagnostic {
    slotKey: string;
    slotLabel: string;
    bindingType: "bound" | "title-related" | "current-related" | "unconfigured";
    rawExpression: string;
    warning?: string;
}

export interface InfoCardProps {
    data: InfoCardData;
    layout: LayoutMode;
    hideEmpty: boolean;
    showBorder: boolean;
    showVersion: boolean;
    showTitle: boolean;
    startExpanded: boolean;
    /** Separator string rendered between subtitle parts. Default "·" (middle dot). */
    subtitleSeparator?: string;
    /** Optional muted-colour prefix rendered before the title (e.g. "Case: ", "Work Order: "). */
    titlePrefix?: string;
    /** Avatar / image shape. Default "rounded" (rounded rect). */
    imageShape?: "rounded" | "circle" | "square";
    /** Which sections collapse when the chevron is tapped. Default "body" (details + grid). "none" disables collapse entirely. */
    collapsibleSections?: "none" | "body" | "body-tags" | "all";
    /** Show auto-detected leading icons on Smart/Contact detail rows. Default true. No effect on Compact. */
    showDetailIcons?: boolean;
    /** How to render the field's display name on Smart/Contact detail rows. Default "none". No effect on Compact. */
    detailLabelStyle?: "none" | "inline-bold" | "above";
    /** Form factor from context.client.getFormFactor(): 0=mobile, 1=tablet, 2=web. Used to gate desktop-only affordances. */
    formFactor?: number;
    designTime?: boolean;
    theme: InfoCardTheme;
    version: string;
    relatedMappings: RelatedFieldMapping[];
    currentRecordMappings?: RelatedFieldMapping[];
    currentRecordEntityType?: string;
    currentRecordId?: string;
    fetchRelatedData?: (entityType: string, id: string, columns: string[]) => Promise<Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string; color?: string }>>;
    /** Resolves labels, formatted values, and colors for bound fields via record fetch + metadata */
    resolveRecordFields?: () => Promise<Record<string, { label: string; value: string; color?: string }>>;
    onOpenRecord?: (entityType: string, id: string) => void;
    bindingDiagnostics?: BindingDiagnostic[];
    /**
     * Localized strings. Optional — when omitted, English defaults from
     * DEFAULT_STRINGS are used. Production code (index.ts) builds this via
     * context.resources.getString() so the user's Dataverse UI language wins.
     */
    strings?: Partial<InfoCardStrings>;
}

/**
 * Localized strings consumed by the React layouts and helpers. Keys mirror the
 * resx file under strings/SampleInfoCard.<LCID>.resx. To add a translation,
 * copy SampleInfoCard.1033.resx, rename to the target LCID, translate values,
 * and register it in ControlManifest.Input.xml under <resources>.
 */
export interface InfoCardStrings {
    sectionContact: string;
    sectionDetails: string;
    sectionInfo: string;
    /** Phone link aria-label/title. {0} = phone number */
    actionCall: string;
    /** Email link aria-label/title. {0} = email address */
    actionEmail: string;
    /** Map link aria-label/title. {0} = address text */
    actionOpenInMaps: string;
    /** Web link aria-label/title. {0} = URL */
    actionOpenWebsite: string;
    /** Lookup row aria-label. {0} = record name */
    actionOpenRecord: string;
    cardExpand: string;
    cardCollapse: string;
    durationDaysSuffix: string;
    durationHoursSuffix: string;
    durationMinutesSuffix: string;
    durationZero: string;
    /** Copy-to-clipboard button aria-label/title. {0} = value being copied */
    actionCopy: string;
    /** Live-region announcement after a successful copy */
    actionCopied: string;
}

/**
 * English defaults used when the host doesn't provide localized strings (e.g.
 * jest/jsdom tests, unconfigured harness). Production runs always pass a
 * resolved strings bag from index.ts via context.resources.getString().
 */
export const DEFAULT_STRINGS: InfoCardStrings = {
    sectionContact: "Contact",
    sectionDetails: "Details",
    sectionInfo: "Info",
    actionCall: "Call {0}",
    actionEmail: "Email {0}",
    actionOpenInMaps: "Open in Maps: {0}",
    actionOpenWebsite: "Open website {0}",
    actionOpenRecord: "Open record {0}",
    cardExpand: "Expand card",
    cardCollapse: "Collapse card",
    durationDaysSuffix: "d",
    durationHoursSuffix: "h",
    durationMinutesSuffix: "m",
    durationZero: "0m",
    actionCopy: "Copy {0}",
    actionCopied: "Copied",
};

/** Format a localized template string. Replaces {0} with `arg`. */
export function formatTemplate(template: string, arg: string): string {
    return template.replace("{0}", arg);
}

/**
 * Format a duration in minutes using localized unit suffixes. Reused by both
 * the read-time formatter (index.ts readSlot) and any consumer of `strings`.
 *
 * Negative inputs return their string representation unchanged (matches the
 * pre-localization behavior used by existing tests).
 */
export function formatLocalizedDuration(
    minutes: number,
    strings: InfoCardStrings,
    formatInteger?: (n: number) => string,
): string {
    if (minutes < 0) return String(minutes);
    const fmt = (n: number): string => {
        if (!formatInteger) return String(n);
        try { return formatInteger(n); } catch { return String(n); }
    };
    const days = Math.floor(minutes / 1440);
    const hrs = Math.floor((minutes % 1440) / 60);
    const mins = minutes % 60;
    const parts: string[] = [];
    if (days > 0) parts.push(`${fmt(days)}${strings.durationDaysSuffix}`);
    if (hrs > 0) parts.push(`${fmt(hrs)}${strings.durationHoursSuffix}`);
    if (mins > 0) parts.push(`${fmt(mins)}${strings.durationMinutesSuffix}`);
    return parts.length > 0 ? parts.join(" ") : strings.durationZero;
}

// ════════════════════════════════════════════════════════════════════
// Inline SVG Icons
// ════════════════════════════════════════════════════════════════════

interface IconProps {
    size?: number;
    color?: string;
}

const PhoneIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M3.6 1.5C3.2 1.1 2.5 1.1 2.1 1.5L1.2 2.4C0.5 3.1 0.3 4.1 0.7 5c1.5 3.3 3.5 6 6.3 8.3 0.9 0.6 1.9 0.4 2.6-0.3l0.9-0.9c0.4-0.4 0.4-1.1 0-1.5l-2-2c-0.4-0.4-1-0.4-1.4 0L6.4 9.3c-0.1 0.1-0.3 0.1-0.4 0C4.5 8.1 3.4 6.9 2.7 5.5c-0.1-0.1-0.1-0.3 0-0.4l0.9-0.9c0.4-0.4 0.4-1 0-1.4L3.6 1.5z" fill={color} />
    </svg>
);

const MobileIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M5 1C4.4 1 4 1.4 4 2v12c0 0.6 0.4 1 1 1h6c0.6 0 1-0.4 1-1V2c0-0.6-0.4-1-1-1H5zm0 1h6v10H5V2zm2 11h2v1H7v-1z" fill={color} />
    </svg>
);

const EmailIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M1 4.5C1 3.7 1.7 3 2.5 3h11C14.3 3 15 3.7 15 4.5v7c0 0.8-0.7 1.5-1.5 1.5h-11C1.7 13 1 12.3 1 11.5v-7zm1.2 0.3v0.4L8 8.5l5.8-3.3v-0.4c0-0.2-0.1-0.3-0.3-0.3h-11c-0.2 0-0.3 0.1-0.3 0.3zm11.6 1.5L8.3 9.4c-0.2 0.1-0.4 0.1-0.6 0L2.2 6.3v5.2c0 0.2 0.1 0.3 0.3 0.3h11c0.2 0 0.3-0.1 0.3-0.3V6.3z" fill={color} />
    </svg>
);

const WebIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <circle cx="8" cy="8" r="6.5" stroke={color} strokeWidth="1.2" fill="none" />
        <ellipse cx="8" cy="8" rx="3" ry="6.5" stroke={color} strokeWidth="1" fill="none" />
        <line x1="1.5" y1="8" x2="14.5" y2="8" stroke={color} strokeWidth="1" />
    </svg>
);

const PinIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M8 1C5.2 1 3 3.2 3 6c0 3.5 4.1 8.3 4.7 9 0.2 0.2 0.4 0.2 0.6 0C8.9 14.3 13 9.5 13 6c0-2.8-2.2-5-5-5zm0 7c-1.1 0-2-0.9-2-2s0.9-2 2-2 2 0.9 2 2-0.9 2-2 2z" fill={color} />
    </svg>
);

const InfoIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M8 1a7 7 0 100 14A7 7 0 008 1zm0 3a0.8 0.8 0 110 1.6A0.8 0.8 0 018 4zm1 8H7v-1h0.5V7.5H7v-1h1.5V11H9v1z" fill={color} />
    </svg>
);

const CalendarIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M5 1v1H3.5C2.7 2 2 2.7 2 3.5v10C2 14.3 2.7 15 3.5 15h9c0.8 0 1.5-0.7 1.5-1.5v-10C14 2.7 13.3 2 12.5 2H11V1h-1v1H6V1H5zM3 5h10v8.5c0 0.3-0.2 0.5-0.5 0.5h-9C3.2 14 3 13.8 3 13.5V5zm1 1.5v2h2v-2H4zm3 0v2h2v-2H7zm3 0v2h2v-2h-2zm-6 3v2h2v-2H4zm3 0v2h2v-2H7z" fill={color} />
    </svg>
);

const GearIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M7.1 1h1.8l0.3 1.8c0.4 0.1 0.8 0.3 1.2 0.6l1.7-0.7 0.9 1.6-1.4 1.1c0.1 0.4 0.1 0.8 0 1.2l1.4 1.1-0.9 1.6-1.7-0.7c-0.4 0.3-0.8 0.5-1.2 0.6L8.9 15H7.1l-0.3-1.8c-0.4-0.1-0.8-0.3-1.2-0.6l-1.7 0.7-0.9-1.6 1.4-1.1c-0.1-0.4-0.1-0.8 0-1.2L3 8.3l0.9-1.6 1.7 0.7c0.4-0.3 0.8-0.5 1.2-0.6L7.1 1zM8 5.5a2.5 2.5 0 100 5 2.5 2.5 0 000-5z" fill={color} />
    </svg>
);

const PersonIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill={color} xmlns="http://www.w3.org/2000/svg">
        <circle cx="8" cy="4.5" r="2.5" />
        <path d="M3 14.5c0-2.8 2.2-5 5-5s5 2.2 5 5H3z" />
    </svg>
);

const ChevronDown: React.FC<IconProps> = ({ size = 12, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M2.5 4.5L6 8l3.5-3.5" stroke={color} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
);

// Entity-specific icons for subtitle lookups
const BuildingIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M3 2a1 1 0 011-1h8a1 1 0 011 1v12h1v1H2v-1h1V2zm2 1v1h2V3H5zm4 0v1h2V3H9zM5 5.5v1h2v-1H5zm4 0v1h2v-1H9zM5 8v1h2V8H5zm4 0v1h2V8H9zM5 10.5v1h2v-1H5zm4 0v1h2v-1H9z" fill={color} />
    </svg>
);

const ClipboardIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M6 1a1 1 0 00-1 1H4a1 1 0 00-1 1v11a1 1 0 001 1h8a1 1 0 001-1V3a1 1 0 00-1-1h-1a1 1 0 00-1-1H6zm0 1h4v1H6V2zM5 6h6v1H5V6zm0 2.5h6v1H5v-1zm0 2.5h4v1H5v-1z" fill={color} />
    </svg>
);

const TicketIcon: React.FC<IconProps> = ({ size = 14, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M2 3a1 1 0 011-1h10a1 1 0 011 1v3.5a1.5 1.5 0 100 3V13a1 1 0 01-1 1H3a1 1 0 01-1-1V9.5a1.5 1.5 0 100-3V3zm3 2v1h6V5H5zm0 3v1h4V8H5z" fill={color} />
    </svg>
);

/** Map known Dataverse entity types to icons. Returns null for unknown entities. */
function entityIcon(entityType: string | undefined, size: number, color: string): React.ReactElement | null {
    if (!entityType) return null;
    const et = entityType.toLowerCase();
    if (et === "account") return <BuildingIcon size={size} color={color} />;
    if (et === "contact") return <PersonIcon size={size} color={color} />;
    if (et === "bookableresource") return <PersonIcon size={size} color={color} />;
    if (et === "msdyn_workorder") return <ClipboardIcon size={size} color={color} />;
    if (et === "incident") return <TicketIcon size={size} color={color} />;
    if (et === "systemuser") return <PersonIcon size={size} color={color} />;
    return null;
}

// ════════════════════════════════════════════════════════════════════
// Helpers
// ════════════════════════════════════════════════════════════════════

function buildMapUrl(lat: number | null, lng: number | null): string | null {
    if (lat == null || lng == null) return null;
    if (lat === 0 && lng === 0) return null;
    return `https://maps.google.com/maps?q=${lat},${lng}`;
}

function filterEmpty(fields: SlotField[], hideEmpty: boolean): SlotField[] {
    if (!hideEmpty) return fields;
    return fields.filter(f => !f.isEmpty || f.isPending);
}

// ── Loading affordances: shimmer + top progress bar ──────────────
// Inline keyframes (styles are otherwise inline; one global <style> is
// the simplest way to drive CSS animations).
const LOADING_KEYFRAMES = `
@keyframes infocard-shimmer {
  0% { background-position: -200px 0; }
  100% { background-position: calc(200px + 100%) 0; }
}
@keyframes infocard-progress {
  0% { transform: translateX(-100%); }
  100% { transform: translateX(400%); }
}`;

const LoadingKeyframes: React.FC = () => (
    <style dangerouslySetInnerHTML={{ __html: LOADING_KEYFRAMES }} />
);

interface ShimmerProps {
    theme: InfoCardTheme;
    width?: string | number;
    height?: string | number;
}

const Shimmer: React.FC<ShimmerProps> = ({ theme, width = "70%", height = "0.9em" }) => (
    <span
        aria-hidden="true"
        style={{
            display: "inline-block",
            verticalAlign: "middle",
            width: typeof width === "number" ? `${width}px` : width,
            height: typeof height === "number" ? `${height}px` : height,
            borderRadius: 4,
            background: `linear-gradient(90deg, ${theme.borderLight} 0%, ${theme.border} 50%, ${theme.borderLight} 100%)`,
            backgroundSize: "200px 100%",
            backgroundRepeat: "no-repeat",
            animation: "infocard-shimmer 1.4s linear infinite",
            opacity: 0.85,
        }}
    />
);

const TopProgressBar: React.FC<{ theme: InfoCardTheme }> = ({ theme }) => (
    <div
        aria-hidden="true"
        role="progressbar"
        style={{
            position: "absolute",
            top: 0,
            left: 0,
            right: 0,
            height: 2,
            overflow: "hidden",
            borderTopLeftRadius: theme.radius,
            borderTopRightRadius: theme.radius,
            background: theme.borderLight,
            pointerEvents: "none",
            zIndex: 1,
        }}
    >
        <div
            style={{
                position: "absolute",
                top: 0,
                left: 0,
                height: "100%",
                width: "25%",
                background: theme.brand,
                animation: "infocard-progress 1.6s ease-in-out infinite",
                borderRadius: 2,
            }}
        />
    </div>
);

/** Renders a shimmer when `field.isPending`, otherwise the supplied children (defaults to field.value).
 *  When the field has both dateText and timeText (DateAndTime.DateAndTime), they are stacked on two lines
 *  unless the caller passes children to override. */
function ValueOrShimmer(props: {
    field: SlotField;
    theme: InfoCardTheme;
    width?: string | number;
    children?: React.ReactNode;
}): React.ReactElement {
    if (props.field.isPending) {
        return <Shimmer theme={props.theme} width={props.width ?? "70%"} />;
    }
    if (props.children !== undefined) {
        return <>{props.children}</>;
    }
    const { dateText, timeText } = props.field;
    if (dateText && timeText) {
        return (
            <span style={{ display: "inline-flex", flexDirection: "column", lineHeight: 1.2 }}>
                <span>{dateText}</span>
                <span style={{ color: props.theme.textMuted, fontSize: "0.9em" }}>{timeText}</span>
            </span>
        );
    }
    return <>{props.field.value}</>;
}

// ════════════════════════════════════════════════════════════════════
// Copy-to-clipboard (desktop only)
// ════════════════════════════════════════════════════════════════════

/**
 * Returns a stable copy callback and a transient "copied" announcement.
 * Hook is no-op-safe when navigator.clipboard is unavailable (returns null cb).
 *
 * The "copied" string is exposed so callers can render it into an aria-live
 * region for screen-reader users; it auto-clears after 1500ms so re-copies
 * still announce.
 */
function useClipboardCopy(announceText: string): {
    copy: ((value: string) => void) | null;
    announcement: string;
} {
    const [announcement, setAnnouncement] = React.useState("");
    const timerRef = React.useRef<number | null>(null);

    const copy = React.useMemo(() => {
        if (typeof navigator === "undefined" || !navigator.clipboard?.writeText) {
            return null;
        }
        return (value: string) => {
            navigator.clipboard.writeText(value).then(
                () => {
                    setAnnouncement("");
                    // Force re-trigger so consecutive copies still announce.
                    requestAnimationFrame(() => setAnnouncement(announceText));
                    if (timerRef.current) window.clearTimeout(timerRef.current);
                    timerRef.current = window.setTimeout(() => setAnnouncement(""), 1500);
                },
                (err) => console.warn("[InfoCard] clipboard copy failed", err),
            );
        };
    }, [announceText]);

    React.useEffect(() => () => {
        if (timerRef.current) window.clearTimeout(timerRef.current);
    }, []);

    return { copy, announcement };
}

interface CopyButtonProps {
    value: string;
    label: string;
    theme: InfoCardTheme;
    onCopy: (value: string) => void;
}

/** Small icon-only button that copies a value. Renders nothing when value is empty. */
const CopyButton: React.FC<CopyButtonProps> = ({ value, label, theme, onCopy }) => {
    if (!value) return null;
    return (
        <button
            type="button"
            onClick={(e) => { e.stopPropagation(); e.preventDefault(); onCopy(value); }}
            aria-label={label}
            title={label}
            style={{
                display: "inline-flex",
                alignItems: "center",
                justifyContent: "center",
                width: 24,
                height: 24,
                marginLeft: 4,
                padding: 0,
                border: "none",
                background: "transparent",
                color: theme.textMuted,
                cursor: "pointer",
                borderRadius: 4,
            }}
        >
            <ClipboardIcon size={14} color="currentColor" />
        </button>
    );
};

function guessDetailIcon(field: SlotField, color: string): React.ReactElement | null {
    const lbl = field.label.toLowerCase();
    if (lbl.includes("address") || lbl.includes("location")) {
        return <PinIcon size={14} color={color} />;
    }
    if (lbl.includes("phone") || lbl.includes("tel")) {
        return <PhoneIcon size={14} color={color} />;
    }
    if (lbl.includes("email")) {
        return <EmailIcon size={14} color={color} />;
    }
    if (lbl.includes("instruction") || lbl.includes("note") || lbl.includes("summary") || lbl.includes("description")) {
        return <InfoIcon size={14} color={color} />;
    }
    if (lbl.includes("date") || lbl.includes("time") || lbl.includes("schedule")) {
        return <CalendarIcon size={14} color={color} />;
    }
    if (lbl.includes("serial") || lbl.includes("device") || lbl.includes("asset") || lbl.includes("tag")) {
        return <GearIcon size={14} color={color} />;
    }
    return null;
}

function deduplicateSubtitles(subtitles: SlotField[]): SlotField[] {
    const seen = new Set<string>();
    const result: SlotField[] = [];
    for (const s of subtitles) {
        const key = s.value.toLowerCase().trim();
        if (seen.has(key)) continue;
        seen.add(key);
        result.push(s);
    }
    return result;
}

// ════════════════════════════════════════════════════════════════════
// mergeRelatedFields
// ════════════════════════════════════════════════════════════════════

export function mergeRelatedFields(
    data: InfoCardData,
    relatedFields: Record<string, SlotField>,
    mappings: RelatedFieldMapping[],
): InfoCardData {
    const merged: InfoCardData = {
        ...data,
        subtitles: [...data.subtitles],
        phones: [...data.phones],
        details: [...data.details],
        gridFields: [...data.gridFields],
        tags: [...data.tags],
    };

    // Replace by slotKey rather than digit-derived index. readSlotGroup compacts unconfigured
    // slots out of the array, so the trailing digit (e.g. subtitleField2 → 1) does not match
    // the post-compaction position. Fall back to position-from-digit only when the array does
    // not yet contain a placeholder for this slot (e.g. readSlot returned null entirely).
    function placeBySlotKey(
        arr: SlotField[],
        target: string,
        groupPrefix: string,
        fetched: SlotField,
    ): void {
        const existingIdx = arr.findIndex(f => f.slotKey === target);
        if (existingIdx >= 0) {
            arr[existingIdx] = fetched;
            return;
        }
        const digit = parseInt(target.replace(groupPrefix, ""), 10);
        if (!Number.isNaN(digit)) {
            const idx = digit - 1;
            while (arr.length < idx) {
                arr.push({
                    slotKey: `${groupPrefix}${arr.length + 1}`,
                    label: "", value: "---", rawValue: null, isEmpty: true,
                });
            }
            arr.splice(idx, 0, fetched);
            return;
        }
        arr.push(fetched);
    }

    for (const mapping of mappings) {
        const fetched = relatedFields[mapping.fetchField];
        if (!fetched) continue;

        const target = mapping.targetSlot;
        // Stamp slotKey so the override pass can match by slot name regardless of array-index drift
        // when readSlotGroup compacts unconfigured slots.
        const fetchedWithKey: SlotField = { ...fetched, slotKey: target };

        if (target === "addressField") {
            merged.address = fetchedWithKey;
        } else if (target === "imageField") {
            const url = (typeof fetchedWithKey.rawValue === "string" && fetchedWithKey.rawValue.length > 0)
                ? fetchedWithKey.rawValue
                : (typeof fetchedWithKey.value === "string" && fetchedWithKey.value.length > 0 && fetchedWithKey.value !== "---")
                    ? fetchedWithKey.value
                    : null;
            if (url) merged.imageUrl = url;
        } else if (target === "phoneField1" || target === "phoneField2") {
            placeBySlotKey(merged.phones, target, "phoneField", fetchedWithKey);
        } else if (target === "emailField") {
            merged.email = fetchedWithKey;
        } else if (target === "webField") {
            merged.web = fetchedWithKey;
        } else if (target.startsWith("subtitleField")) {
            placeBySlotKey(merged.subtitles, target, "subtitleField", fetchedWithKey);
        } else if (target.startsWith("detailField")) {
            placeBySlotKey(merged.details, target, "detailField", fetchedWithKey);
        } else if (target.startsWith("gridField")) {
            placeBySlotKey(merged.gridFields, target, "gridField", fetchedWithKey);
        } else if (target.startsWith("tagField")) {
            placeBySlotKey(merged.tags, target, "tagField", fetchedWithKey);
        }
    }

    return merged;
}

// ════════════════════════════════════════════════════════════════════
// Sub-layout components
// ════════════════════════════════════════════════════════════════════

// ── Shared: Title + Subtitles ─────────────────────────────────────

interface HeaderProps {
    data: InfoCardData;
    theme: InfoCardTheme;
    hideEmpty: boolean;
    showTitle?: boolean;
    designTime?: boolean;
    onOpenRecord?: (entityType: string, id: string) => void;
    strings: InfoCardStrings;
    /** Separator string rendered between subtitle parts. Default "·" (middle dot). */
    subtitleSeparator?: string;
    /** Muted-colour prefix rendered immediately before the title text (e.g. "Case: "). */
    titlePrefix?: string;
}

const Header: React.FC<HeaderProps> = ({ data, theme, hideEmpty, showTitle = true, designTime, onOpenRecord, strings, subtitleSeparator, titlePrefix }) => {
    const title = data.title;
    if (!title || (!designTime && title.isEmpty)) return null;

    const isLookup = !!(title.lookupEntityType && title.lookupId);
    const titleCanOpen = isLookup && !!onOpenRecord;
    const subtitles = deduplicateSubtitles(
        filterEmpty(data.subtitles, designTime ? false : hideEmpty),
    );

    return (
        <div style={{ marginBottom: 8 }}>
            {showTitle && (
                <div
                    style={{
                        fontSize: 16,
                        fontWeight: 600,
                        color: isLookup ? theme.brand : theme.textPrimary,
                        cursor: titleCanOpen ? "pointer" : "default",
                        lineHeight: "22px",
                    }}
                    onClick={titleCanOpen ? (e) => { e.stopPropagation(); onOpenRecord!(title.lookupEntityType!, title.lookupId!); } : undefined}
                    onKeyDown={titleCanOpen ? (e) => {
                        if (e.key === "Enter" || e.key === " ") {
                            e.preventDefault();
                            e.stopPropagation();
                            onOpenRecord!(title.lookupEntityType!, title.lookupId!);
                        }
                    } : undefined}
                    role={titleCanOpen ? "button" : undefined}
                    tabIndex={titleCanOpen ? 0 : undefined}
                    aria-label={titleCanOpen ? formatTemplate(strings.actionOpenRecord, String(title.value)) : undefined}
                    title={title.label}
                >
                    {titlePrefix && (
                        <span style={{ color: theme.textMuted, fontWeight: 500 }}>{titlePrefix}</span>
                    )}
                    {title.value}
                </div>
            )}
            {subtitles.length > 0 && (
                <div style={{ fontSize: 13, color: theme.textSecondary, lineHeight: "18px", marginTop: showTitle ? 2 : 0 }}>
                    {subtitles.map((sub, i) => {
                        const isSubLookup = !!(sub.lookupEntityType && sub.lookupId);
                        const subCanOpen = isSubLookup && !!onOpenRecord;
                        const icon = entityIcon(sub.lookupEntityType, 12, theme.textSecondary);
                        return (
                            <React.Fragment key={i}>
                                {i > 0 && (
                                    <span style={{ margin: "0 6px", color: theme.textMuted }}>{subtitleSeparator || "\u00b7"}</span>
                                )}
                                {icon && <span style={{ marginRight: 3, verticalAlign: "middle", display: "inline-flex" }} aria-hidden="true">{icon}</span>}
                                <span
                                    style={{
                                        color: isSubLookup ? theme.brand : theme.textSecondary,
                                        cursor: subCanOpen ? "pointer" : "default",
                                    }}
                                    onClick={
                                        subCanOpen
                                            ? (e) => { e.stopPropagation(); onOpenRecord!(sub.lookupEntityType!, sub.lookupId!); }
                                            : undefined
                                    }
                                    onKeyDown={subCanOpen ? (e) => {
                                        if (e.key === "Enter" || e.key === " ") {
                                            e.preventDefault();
                                            e.stopPropagation();
                                            onOpenRecord!(sub.lookupEntityType!, sub.lookupId!);
                                        }
                                    } : undefined}
                                    role={subCanOpen ? "button" : undefined}
                                    tabIndex={subCanOpen ? 0 : undefined}
                                    aria-label={subCanOpen ? formatTemplate(strings.actionOpenRecord, String(sub.value)) : undefined}
                                    title={sub.label}
                                >
                                    <ValueOrShimmer field={sub} theme={theme} width="80px" />
                                </span>
                            </React.Fragment>
                        );
                    })}
                </div>
            )}
        </div>
    );
};

// ── Shared: Contact rows (address, phones, email, web) ────────────

interface ContactRowsProps {
    data: InfoCardData;
    theme: InfoCardTheme;
    hideEmpty: boolean;
    strings: InfoCardStrings;
    /** 0=mobile, 1=tablet, 2=web. Copy-to-clipboard buttons render only when 2 (desktop). */
    formFactor?: number;
}

const ContactRows: React.FC<ContactRowsProps> = ({ data, theme, hideEmpty, strings, formFactor }) => {
    const mapUrl = buildMapUrl(data.latitude, data.longitude);
    const address = data.address;
    const phones = filterEmpty(data.phones, hideEmpty);
    const email = data.email && (!data.email.isEmpty || data.email.isPending) ? data.email : null;
    const web = data.web && (!data.web.isEmpty || data.web.isPending) ? data.web : null;

    const hasAny = (address && (!address.isEmpty || address.isPending)) || phones.length > 0 || email || web;
    const { copy, announcement } = useClipboardCopy(strings.actionCopied);
    if (!hasAny) return null;
    const showCopy = formFactor === 2 && !!copy;
    const renderCopy = (value: string) => (
        showCopy && copy
            ? <CopyButton value={value} label={formatTemplate(strings.actionCopy, value)} theme={theme} onCopy={copy} />
            : null
    );
    const groupStyle: React.CSSProperties = { display: "inline-flex", alignItems: "center" };

    const rowStyle: React.CSSProperties = {
        display: "flex",
        alignItems: "flex-start",
        gap: 8,
        padding: "4px 0",
        fontSize: 13,
        color: theme.textSecondary,
        lineHeight: "18px",
    };

    const chipStyle: React.CSSProperties = {
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        padding: "6px 10px",
        fontSize: 13,
        color: theme.textPrimary,
        textDecoration: "none",
        borderRadius: 6,
        cursor: "pointer",
        minWidth: 0,
    };

    return (
        <div style={{
            display: "flex",
            flexWrap: "wrap",
            gap: 0,
            borderTop: `1px solid ${theme.borderLight}`,
            borderBottom: `1px solid ${theme.borderLight}`,
            padding: "0 6px",
            margin: "8px 0 0",
        }}>
            {/* Address */}
            {address && (!address.isEmpty || address.isPending) && (
                <span style={groupStyle}>
                    {address.isPending ? (
                        <span style={{ ...chipStyle, color: theme.textSecondary, cursor: "default" }} title={address.label}>
                            <PinIcon size={16} color={theme.textMuted} />
                            <Shimmer theme={theme} width="120px" />
                        </span>
                    ) : mapUrl ? (
                        <a
                            href={mapUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            style={chipStyle}
                            title={address.label}
                            aria-label={formatTemplate(strings.actionOpenInMaps, String(address.value))}
                            onClick={(e) => e.stopPropagation()}
                        >
                            <PinIcon size={16} color={theme.brand} />
                            {address.value}
                        </a>
                    ) : (
                        <span style={{ ...chipStyle, color: theme.textSecondary, cursor: "default" }} title={address.label}>
                            <PinIcon size={16} color={theme.textMuted} />
                            {address.value}
                        </span>
                    )}
                    {!address.isPending && renderCopy(String(address.value))}
                </span>
            )}

            {/* Phones — phone1 gets landline icon, phone2 gets mobile icon */}
            {phones.map((phone, i) => (
                <span key={`phone-${i}`} style={groupStyle}>
                    {phone.isPending ? (
                        <span style={{ ...chipStyle, cursor: "default" }} title={phone.label}>
                            {i === 0 ? <PhoneIcon size={14} color={theme.textMuted} /> : <MobileIcon size={14} color={theme.textMuted} />}
                            <Shimmer theme={theme} width="90px" />
                        </span>
                    ) : (
                        <a
                            href={`tel:${String(phone.value).replace(/\s+/g, "")}`}
                            style={chipStyle}
                            title={phone.label}
                            aria-label={formatTemplate(strings.actionCall, String(phone.value))}
                            onClick={(e) => e.stopPropagation()}
                        >
                            {i === 0 ? <PhoneIcon size={14} color={theme.brand} /> : <MobileIcon size={14} color={theme.brand} />}
                            {phone.value}
                        </a>
                    )}
                    {!phone.isPending && renderCopy(String(phone.value))}
                </span>
            ))}

            {/* Email */}
            {email && (
                <span style={groupStyle}>
                    {email.isPending ? (
                        <span style={{ ...chipStyle, cursor: "default" }} title={email.label}>
                            <EmailIcon size={14} color={theme.textMuted} />
                            <Shimmer theme={theme} width="110px" />
                        </span>
                    ) : (
                        <a
                            href={`mailto:${email.value}`}
                            style={chipStyle}
                            title={email.label}
                            aria-label={formatTemplate(strings.actionEmail, String(email.value))}
                            onClick={(e) => e.stopPropagation()}
                        >
                            <EmailIcon size={14} color={theme.brand} />
                            {email.value}
                        </a>
                    )}
                    {!email.isPending && renderCopy(String(email.value))}
                </span>
            )}

            {/* Web */}
            {web && (
                web.isPending ? (
                    <span style={{ ...chipStyle, cursor: "default" }} title={web.label}>
                        <WebIcon size={14} color={theme.textMuted} />
                        <Shimmer theme={theme} width="110px" />
                    </span>
                ) : (() => {
                    const safeWeb = safeHttpUrl(web.value);
                    if (!safeWeb) {
                        return (
                            <span style={{ ...chipStyle, cursor: "default", color: theme.textMuted }} title={web.label}>
                                <WebIcon size={14} color={theme.textMuted} />
                                {web.value}
                            </span>
                        );
                    }
                    return (
                        <a
                            href={safeWeb}
                            target="_blank"
                            rel="noopener noreferrer"
                            style={chipStyle}
                            title={web.label}
                            aria-label={formatTemplate(strings.actionOpenWebsite, String(web.value))}
                            onClick={(e) => e.stopPropagation()}
                        >
                            <WebIcon size={14} color={theme.brand} />
                            {web.value}
                        </a>
                    );
                })()
            )}
            {/* SR-only live region for copy success announcements */}
            <span
                aria-live="polite"
                aria-atomic="true"
                style={{ position: "absolute", width: 1, height: 1, padding: 0, margin: -1, overflow: "hidden", clip: "rect(0,0,0,0)", whiteSpace: "nowrap", border: 0 }}
            >
                {announcement}
            </span>
        </div>
    );
};

// ── Shared: Detail rows ───────────────────────────────────────────

interface DetailRowsProps {
    details: SlotField[];
    theme: InfoCardTheme;
    hideEmpty: boolean;
    latitude: number | null;
    longitude: number | null;
    strings: InfoCardStrings;
    /** Show leading auto-detected icon. Default true. */
    showIcons?: boolean;
    /** How to render the field's display name. Default "none". */
    labelStyle?: "none" | "inline-bold" | "above";
}

const DetailRows: React.FC<DetailRowsProps> = ({ details, theme, hideEmpty, latitude, longitude, strings, showIcons = true, labelStyle = "none" }) => {
    const filtered = filterEmpty(details, hideEmpty);
    if (filtered.length === 0) return null;

    const mapUrl = buildMapUrl(latitude, longitude);

    return (
        <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
            {filtered.map((field, i) => {
                const icon = showIcons ? guessDetailIcon(field, theme.textMuted) : null;
                const isAddressLike = field.label.toLowerCase().includes("address") || field.label.toLowerCase().includes("location");
                const showMapLink = isAddressLike && mapUrl;
                const hasLabel = !!field.label && labelStyle !== "none";

                const valueNode = showMapLink ? (
                    <a
                        href={mapUrl!}
                        target="_blank"
                        rel="noopener noreferrer"
                        style={{ color: theme.brand, textDecoration: "none", cursor: "pointer" }}
                        aria-label={formatTemplate(strings.actionOpenInMaps, String(field.value))}
                        onClick={(e) => e.stopPropagation()}
                    >
                        <ValueOrShimmer field={field} theme={theme} width="60%" />
                    </a>
                ) : (
                    <ValueOrShimmer field={field} theme={theme} width="80%" />
                );

                if (labelStyle === "above") {
                    return (
                        <div
                            key={i}
                            style={{
                                display: "flex",
                                alignItems: "flex-start",
                                gap: 8,
                                padding: "4px 0",
                                fontSize: 13,
                                color: theme.textSecondary,
                                lineHeight: "18px",
                            }}
                            title={field.label}
                        >
                            {icon && (
                                <span style={{ flexShrink: 0, marginTop: 2 }}>{icon}</span>
                            )}
                            <div style={{ flex: 1, minWidth: 0 }}>
                                {hasLabel && (
                                    <div
                                        style={{
                                            fontSize: 11,
                                            fontWeight: 600,
                                            textTransform: "uppercase",
                                            letterSpacing: "0.04em",
                                            color: theme.textMuted,
                                            marginBottom: 2,
                                        }}
                                    >
                                        {field.label}
                                    </div>
                                )}
                                <span style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                                    {valueNode}
                                </span>
                            </div>
                        </div>
                    );
                }

                return (
                    <div
                        key={i}
                        style={{
                            display: "flex",
                            alignItems: "flex-start",
                            gap: 8,
                            padding: "4px 0",
                            fontSize: 13,
                            color: theme.textSecondary,
                            lineHeight: "18px",
                        }}
                        title={field.label}
                    >
                        {icon && (
                            <span style={{ flexShrink: 0, marginTop: 2 }}>{icon}</span>
                        )}
                        {showMapLink ? (
                            <span style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                                {labelStyle === "inline-bold" && hasLabel && (
                                    <span style={{ fontWeight: 600, color: theme.textPrimary }}>
                                        {field.label}:{" "}
                                    </span>
                                )}
                                {valueNode}
                            </span>
                        ) : (
                            <span style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                                {labelStyle === "inline-bold" && hasLabel && (
                                    <span style={{ fontWeight: 600, color: theme.textPrimary }}>
                                        {field.label}:{" "}
                                    </span>
                                )}
                                {valueNode}
                            </span>
                        )}
                    </div>
                );
            })}
        </div>
    );
};

// ── Shared: Grid fields ───────────────────────────────────────────

interface GridFieldsProps {
    fields: SlotField[];
    theme: InfoCardTheme;
    hideEmpty: boolean;
}

const GridFields: React.FC<GridFieldsProps> = ({ fields, theme, hideEmpty }) => {
    const filtered = filterEmpty(fields, hideEmpty);
    if (filtered.length === 0) return null;

    const columns = filtered.length === 1 ? "1fr" : "1fr 1fr";

    return (
        <div
            style={{
                display: "grid",
                gridTemplateColumns: columns,
                gap: "8px 16px",
            }}
        >
            {filtered.map((field, i) => (
                <div key={i} style={{ minWidth: 0 }}>
                    <div style={{ fontSize: 11, color: theme.textMuted, lineHeight: "14px", marginBottom: 2 }}>
                        {field.label}
                    </div>
                    <div
                        style={{
                            fontSize: 13,
                            color: theme.textPrimary,
                            lineHeight: "18px",
                            whiteSpace: "normal",
                            wordBreak: "break-word",
                            overflowWrap: "anywhere",
                        }}
                        title={field.value}
                    >
                        <ValueOrShimmer field={field} theme={theme} width="80%" />
                    </div>
                </div>
            ))}
        </div>
    );
};

// ── Shared: Tags ──────────────────────────────────────────────────

interface TagsProps {
    tags: SlotField[];
    theme: InfoCardTheme;
    hideEmpty: boolean;
}

const Tags: React.FC<TagsProps> = ({ tags, theme, hideEmpty }) => {
    const filtered = filterEmpty(tags, hideEmpty);
    if (filtered.length === 0) return null;

    return (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
            {filtered.map((tag, i) => {
                // Use OptionSet color if available, otherwise default brand blue
                const hasColor = !!tag.optionColor;
                const textColor = hasColor ? tag.optionColor! : theme.brand;
                // Lighten the option color for background (10% opacity)
                const bgColor = hasColor ? `${tag.optionColor}1a` : theme.brandLight;
                return (
                    <span
                        key={i}
                        style={{
                            display: "inline-block",
                            padding: "2px 8px",
                            fontSize: 11,
                            fontWeight: 500,
                            color: textColor,
                            background: bgColor,
                            borderRadius: "10px",
                            lineHeight: "16px",
                            whiteSpace: "nowrap",
                        }}
                        title={tag.label}
                    >
                        <ValueOrShimmer field={tag} theme={theme} width="50px" />
                    </span>
                );
            })}
        </div>
    );
};

// ── Shared: Image ─────────────────────────────────────────────────

interface ImageProps {
    imageUrl: string | null;
    title: string;
    theme: InfoCardTheme;
    showInitialsFallback?: boolean;
    /** Avatar shape. Default "rounded" (rounded rect). */
    shape?: "rounded" | "circle" | "square";
}

/** Border radius for the avatar container per shape preference. */
function avatarBorderRadius(shape: "rounded" | "circle" | "square" | undefined): string | number {
    switch (shape) {
        case "circle": return "50%";
        case "square": return 0;
        case "rounded":
        default: return 8;
    }
}

function getInitials(title: string): string {
    if (!title) return "";
    const parts = title.trim().split(/\s+/).filter(Boolean);
    if (parts.length === 0) return "";
    if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

const ImageAvatar: React.FC<ImageProps> = ({ imageUrl, title, theme, showInitialsFallback, shape }) => {
    const [imgError, setImgError] = React.useState(false);
    const onError = React.useCallback(() => setImgError(true), []);
    React.useEffect(() => { setImgError(false); }, [imageUrl]);

    const showImage = !!imageUrl && !imgError;
    const initials = getInitials(title);
    // Only show initials fallback when an image URL was supplied but failed to load.
    // If no image URL at all, render nothing — don't fabricate a placeholder.
    const showFallback = !!imageUrl && imgError && !!showInitialsFallback && initials.length > 0;

    if (!showImage && !showFallback) return null;

    const baseStyle: React.CSSProperties = {
        width: 40,
        height: 40,
        borderRadius: avatarBorderRadius(shape),
        flexShrink: 0,
    };

    if (showImage) {
        return (
            <img
                src={imageUrl as string}
                alt={title || ""}
                onError={onError}
                style={{ ...baseStyle, objectFit: "cover" }}
            />
        );
    }

    return (
        <div
            aria-hidden="true"
            style={{
                ...baseStyle,
                background: theme.brand,
                color: "#FFFFFF",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                fontSize: 14,
                fontWeight: 600,
                letterSpacing: 0.5,
            }}
        >
            {initials}
        </div>
    );
};

// ── Shared: Version badge ─────────────────────────────────────────

interface VersionBadgeProps {
    version: string;
    theme: InfoCardTheme;
}

const VersionBadge: React.FC<VersionBadgeProps> = ({ version, theme }) => {
    const [showTooltip, setShowTooltip] = React.useState(false);

    return (
        <div
            style={{
                position: "absolute",
                bottom: 4,
                right: 4,
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: 4,
            }}
            title={`InfoCard v${version}`}
            onClick={() => setShowTooltip(!showTooltip)}
        >
            {showTooltip && (
                <span
                    style={{
                        fontSize: 10,
                        color: theme.textMuted,
                        background: theme.cardBg,
                        padding: "1px 4px",
                        borderRadius: 3,
                        border: `1px solid ${theme.borderLight}`,
                        whiteSpace: "nowrap",
                    }}
                >
                    v{version}
                </span>
            )}
            <InfoIcon size={12} color={theme.textMuted} />
        </div>
    );
};

// ════════════════════════════════════════════════════════════════════
// Smart Card Layout
// ════════════════════════════════════════════════════════════════════

interface LayoutProps {
    data: InfoCardData;
    theme: InfoCardTheme;
    hideEmpty: boolean;
    showTitle?: boolean;
    startExpanded?: boolean;
    designTime?: boolean;
    onOpenRecord?: (entityType: string, id: string) => void;
    /** Card collapse state, controlled by parent so the whole card surface can toggle. */
    collapsed?: boolean;
    /** True when the card should render the chevron + clickable surface for this layout. */
    isCollapsible?: boolean;
    /** Which sections actually disappear when collapsed. Default "body". */
    collapsibleSections?: "none" | "body" | "body-tags" | "all";
    strings: InfoCardStrings;
    subtitleSeparator?: string;
    /** Muted-colour title prefix passed through to Header. */
    titlePrefix?: string;
    /** Avatar shape passed through to ImageAvatar. */
    imageShape?: "rounded" | "circle" | "square";
    /** Show leading icons on detail rows. Smart/Contact only. */
    showDetailIcons?: boolean;
    /** Render field display name as inline-bold prefix or above-heading on detail rows. Smart/Contact only. */
    detailLabelStyle?: "none" | "inline-bold" | "above";
    /** Form factor: 0=mobile, 1=tablet, 2=web. Plumbed to ContactRows for desktop-only copy buttons. */
    formFactor?: number;
}

/** Returns true when `section` should disappear when the card is collapsed, given the maker setting. */
function shouldCollapseSection(
    setting: "none" | "body" | "body-tags" | "all" | undefined,
    section: "contact" | "body" | "tags",
): boolean {
    switch (setting ?? "body") {
        case "none": return false;
        case "body": return section === "body";
        case "body-tags": return section === "body" || section === "tags";
        case "all": return true;
    }
}

const SmartCardLayout: React.FC<LayoutProps> = ({ data, theme, hideEmpty, showTitle = true, designTime, onOpenRecord, collapsed = false, isCollapsible = false, collapsibleSections, strings, subtitleSeparator, titlePrefix, imageShape, showDetailIcons, detailLabelStyle, formFactor }) => {
    const effectiveHideEmpty = designTime ? false : hideEmpty;
    const details = filterEmpty(data.details, effectiveHideEmpty);
    const gridFields = filterEmpty(data.gridFields, effectiveHideEmpty);
    const tags = filterEmpty(data.tags, effectiveHideEmpty);

    const hasBody = details.length > 0 || gridFields.length > 0;
    const hideContact = collapsed && shouldCollapseSection(collapsibleSections, "contact");
    const hideBody = collapsed && shouldCollapseSection(collapsibleSections, "body");
    const hideTags = collapsed && shouldCollapseSection(collapsibleSections, "tags");

    return (
        <div>
            {/* Header row with image, title, subtitles, and chevron */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                <ImageAvatar imageUrl={data.imageUrl} title={data.title?.value ?? ""} theme={theme} showInitialsFallback={true} shape={imageShape} />
                <div style={{ flex: 1, minWidth: 0 }}>
                    <Header data={data} theme={theme} hideEmpty={hideEmpty} showTitle={showTitle} designTime={designTime} onOpenRecord={onOpenRecord} strings={strings} subtitleSeparator={subtitleSeparator} titlePrefix={titlePrefix} />
                </div>
                {isCollapsible && (
                    <div
                        style={{
                            padding: 4,
                            flexShrink: 0,
                            transition: "transform 0.2s ease",
                            transform: collapsed ? "rotate(-90deg)" : "rotate(0deg)",
                            pointerEvents: "none",
                        }}
                        aria-hidden="true"
                    >
                        <ChevronDown size={12} color={theme.textMuted} />
                    </div>
                )}
            </div>

            {/* Contact section */}
            {!hideContact && (
                <ContactRows data={data} theme={theme} hideEmpty={effectiveHideEmpty} strings={strings} formFactor={formFactor} />
            )}

            {/* Body (details + grid) */}
            {!hideBody && hasBody && (
                <div style={{ marginTop: 8 }}>
                    {details.length > 0 && (
                        <div style={{ marginBottom: gridFields.length > 0 ? 8 : 0 }}>
                            <DetailRows
                                details={data.details}
                                theme={theme}
                                hideEmpty={hideEmpty}
                                latitude={data.latitude}
                                longitude={data.longitude}
                                strings={strings}
                                showIcons={showDetailIcons}
                                labelStyle={detailLabelStyle}
                            />
                        </div>
                    )}
                    {gridFields.length > 0 && (
                        <div style={{ marginTop: details.length > 0 ? 0 : 0 }}>
                            <GridFields fields={data.gridFields} theme={theme} hideEmpty={hideEmpty} />
                        </div>
                    )}
                </div>
            )}

            {/* Tags */}
            {!hideTags && tags.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    <Tags tags={data.tags} theme={theme} hideEmpty={hideEmpty} />
                </div>
            )}
        </div>
    );
};

// ════════════════════════════════════════════════════════════════════
// Contact Card Layout
// ════════════════════════════════════════════════════════════════════

const ContactCardLayout: React.FC<LayoutProps> = ({ data, theme, hideEmpty, showTitle = true, designTime, onOpenRecord, collapsed = false, isCollapsible = false, collapsibleSections, strings, subtitleSeparator, titlePrefix, imageShape, showDetailIcons, detailLabelStyle, formFactor }) => {
    const effectiveHideEmpty = designTime ? false : hideEmpty;
    const details = filterEmpty(data.details, effectiveHideEmpty);
    const gridFields = filterEmpty(data.gridFields, effectiveHideEmpty);
    const tags = filterEmpty(data.tags, effectiveHideEmpty);

    const hideContact = collapsed && shouldCollapseSection(collapsibleSections, "contact");
    const hideBody = collapsed && shouldCollapseSection(collapsibleSections, "body");
    const hideTags = collapsed && shouldCollapseSection(collapsibleSections, "tags");

    return (
        <div>
            {/* Header row with image */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                <ImageAvatar imageUrl={data.imageUrl} title={data.title?.value ?? ""} theme={theme} showInitialsFallback={true} shape={imageShape} />
                <div style={{ flex: 1, minWidth: 0 }}>
                    <Header data={data} theme={theme} hideEmpty={hideEmpty} showTitle={showTitle} designTime={designTime} onOpenRecord={onOpenRecord} strings={strings} subtitleSeparator={subtitleSeparator} titlePrefix={titlePrefix} />
                </div>
                {isCollapsible && (
                    <div
                        style={{
                            padding: 4,
                            flexShrink: 0,
                            transition: "transform 0.2s ease",
                            transform: collapsed ? "rotate(-90deg)" : "rotate(0deg)",
                            pointerEvents: "none",
                        }}
                        aria-hidden="true"
                    >
                        <ChevronDown size={12} color={theme.textMuted} />
                    </div>
                )}
            </div>

            {/* Contact section */}
            {!hideContact && (
                <ContactRows data={data} theme={theme} hideEmpty={effectiveHideEmpty} strings={strings} formFactor={formFactor} />
            )}

            {/* Detail rows */}
            {!hideBody && details.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    <DetailRows
                        details={data.details}
                        theme={theme}
                        hideEmpty={effectiveHideEmpty}
                        latitude={data.latitude}
                        longitude={data.longitude}
                        strings={strings}
                        showIcons={showDetailIcons}
                        labelStyle={detailLabelStyle}
                    />
                </div>
            )}

            {/* Grid fields */}
            {!hideBody && gridFields.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    <GridFields fields={data.gridFields} theme={theme} hideEmpty={hideEmpty} />
                </div>
            )}

            {/* Tags */}
            {!hideTags && tags.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    <Tags tags={data.tags} theme={theme} hideEmpty={hideEmpty} />
                </div>
            )}
        </div>
    );
};

// ════════════════════════════════════════════════════════════════════
// Compact Card Layout
// ════════════════════════════════════════════════════════════════════

const CompactCardLayout: React.FC<LayoutProps> = ({ data, theme, hideEmpty, showTitle = true, designTime, onOpenRecord, collapsed = false, isCollapsible = false, collapsibleSections, strings, subtitleSeparator, titlePrefix, formFactor }) => {
    const effectiveHideEmpty = designTime ? false : hideEmpty;
    const phones = filterEmpty(data.phones, effectiveHideEmpty);
    const email = data.email && (designTime || !data.email.isEmpty || data.email.isPending) ? data.email : null;
    const web = data.web && (designTime || !data.web.isEmpty || data.web.isPending) ? data.web : null;
    const address = data.address && (designTime || !data.address.isEmpty || data.address.isPending) ? data.address : null;
    const details = filterEmpty(data.details, effectiveHideEmpty);
    const gridFields = filterEmpty(data.gridFields, effectiveHideEmpty);
    const tags = filterEmpty(data.tags, effectiveHideEmpty);

    const hasContact = phones.length > 0 || email || web || address;
    const hasDetails = gridFields.length > 0;
    const hasInfo = details.length > 0;

    const hideContact = collapsed && shouldCollapseSection(collapsibleSections, "contact");
    const hideBody = collapsed && shouldCollapseSection(collapsibleSections, "body");
    const hideTags = collapsed && shouldCollapseSection(collapsibleSections, "tags");

    // formFactor not used by Compact today, but kept in signature for parity.
    void formFactor;

    const sectionHeaderStyle: React.CSSProperties = {
        fontSize: 10,
        fontWeight: 600,
        color: theme.textMuted,
        textTransform: "uppercase" as const,
        letterSpacing: "0.5px",
        marginBottom: 4,
        marginTop: 8,
    };

    const fieldRowStyle: React.CSSProperties = {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "baseline",
        padding: "2px 0",
        fontSize: 12,
        gap: 8,
    };

    return (
        <div>
            {/* Header — Compact omits avatar to preserve dense form-feel.
                When collapsible, wrap header + chevron in a flex row. */}
            {isCollapsible ? (
                <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                        <Header data={data} theme={theme} hideEmpty={hideEmpty} showTitle={showTitle} designTime={designTime} onOpenRecord={onOpenRecord} strings={strings} subtitleSeparator={subtitleSeparator} titlePrefix={titlePrefix} />
                    </div>
                    <div
                        style={{
                            padding: 4,
                            flexShrink: 0,
                            transition: "transform 0.2s ease",
                            transform: collapsed ? "rotate(-90deg)" : "rotate(0deg)",
                            pointerEvents: "none",
                        }}
                        aria-hidden="true"
                    >
                        <ChevronDown size={12} color={theme.textMuted} />
                    </div>
                </div>
            ) : (
                <Header data={data} theme={theme} hideEmpty={hideEmpty} showTitle={showTitle} designTime={designTime} onOpenRecord={onOpenRecord} strings={strings} subtitleSeparator={subtitleSeparator} titlePrefix={titlePrefix} />
            )}

            {/* Contact section */}
            {!hideContact && hasContact && (
                <div>
                    <div style={sectionHeaderStyle}>{strings.sectionContact}</div>
                    {address && (
                        <div style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{address.label}</span>
                            <span style={{ color: theme.textPrimary, textAlign: "right" }}>
                                <ValueOrShimmer field={address} theme={theme} width="120px" />
                            </span>
                        </div>
                    )}
                    {phones.map((phone, i) => (
                        <div key={`phone-${i}`} style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{phone.label}</span>
                            {phone.isPending ? (
                                <Shimmer theme={theme} width="100px" />
                            ) : (
                                <a
                                    href={`tel:${String(phone.value).replace(/\s+/g, "")}`}
                                    style={{ color: theme.brand, textDecoration: "none" }}
                                    aria-label={formatTemplate(strings.actionCall, String(phone.value))}
                                >
                                    {phone.value}
                                </a>
                            )}
                        </div>
                    ))}
                    {email && (
                        <div style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{email.label}</span>
                            {email.isPending ? (
                                <Shimmer theme={theme} width="130px" />
                            ) : (
                                <a
                                    href={`mailto:${email.value}`}
                                    style={{ color: theme.brand, textDecoration: "none" }}
                                    aria-label={formatTemplate(strings.actionEmail, String(email.value))}
                                >
                                    {email.value}
                                </a>
                            )}
                        </div>
                    )}
                    {web && (
                        <div style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{web.label}</span>
                            {web.isPending ? (
                                <Shimmer theme={theme} width="130px" />
                            ) : (() => {
                                const safeWeb = safeHttpUrl(web.value);
                                if (!safeWeb) {
                                    return <span style={{ color: theme.textMuted }}>{web.value}</span>;
                                }
                                return (
                                    <a
                                        href={safeWeb}
                                        target="_blank"
                                        rel="noopener noreferrer"
                                        style={{ color: theme.brand, textDecoration: "none" }}
                                        aria-label={formatTemplate(strings.actionOpenWebsite, String(web.value))}
                                    >
                                        {web.value}
                                    </a>
                                );
                            })()}
                        </div>
                    )}
                </div>
            )}

            {/* Details section (grid fields) */}
            {!hideBody && hasDetails && (
                <div>
                    <div style={sectionHeaderStyle}>{strings.sectionDetails}</div>
                    {gridFields.map((field, i) => (
                        <div key={i} style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{field.label}</span>
                            <span style={{ color: theme.textPrimary, textAlign: "right" }}>
                                <ValueOrShimmer field={field} theme={theme} width="100px" />
                            </span>
                        </div>
                    ))}
                </div>
            )}

            {/* Info section (detail fields) */}
            {!hideBody && hasInfo && (
                <div>
                    <div style={sectionHeaderStyle}>{strings.sectionInfo}</div>
                    {details.map((field, i) => (
                        <div key={i} style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{field.label}</span>
                            <span style={{ color: theme.textPrimary, textAlign: "right", wordBreak: "break-word" }}>
                                <ValueOrShimmer field={field} theme={theme} width="140px" />
                            </span>
                        </div>
                    ))}
                </div>
            )}

            {/* Tags */}
            {!hideTags && tags.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    <Tags tags={data.tags} theme={theme} hideEmpty={hideEmpty} />
                </div>
            )}
        </div>
    );
};

// ════════════════════════════════════════════════════════════════════
// Design-Time Binding Panel
// ════════════════════════════════════════════════════════════════════

const DesignTimeBindingPanel: React.FC<{ diagnostics: BindingDiagnostic[]; theme: InfoCardTheme }> = ({ diagnostics, theme }) => {
    const typeColors: Record<string, string> = {
        "bound": "#107c10",
        "title-related": theme.brand,
        "current-related": "#8764b8",
        "unconfigured": theme.textMuted,
    };
    const typeLabels: Record<string, string> = {
        "bound": "Column",
        "title-related": "@ Title",
        "current-related": "@. Record",
        "unconfigured": "---",
    };

    return (
        <div style={{
            marginTop: 10, padding: "8px 10px",
            background: theme.brandLight, borderRadius: theme.radius,
            border: `1px solid ${theme.border}`, fontSize: 11,
        }}>
            <div style={{ fontWeight: 600, fontSize: 12, marginBottom: 6, color: theme.textPrimary }}>
                Slot Bindings
            </div>
            <table style={{ width: "100%", borderCollapse: "collapse" }}>
                <thead>
                    <tr style={{ color: theme.textSecondary, textAlign: "left", borderBottom: `1px solid ${theme.border}` }}>
                        <th style={{ padding: "2px 4px", fontWeight: 500 }}>Slot</th>
                        <th style={{ padding: "2px 4px", fontWeight: 500 }}>Source</th>
                        <th style={{ padding: "2px 4px", fontWeight: 500 }}>Binding</th>
                    </tr>
                </thead>
                <tbody>
                    {diagnostics.map((d) => (
                        <tr key={d.slotKey} style={{
                            borderBottom: `1px solid ${theme.borderLight}`,
                            background: d.warning ? "#fff4ce" : "transparent",
                        }}>
                            <td style={{ padding: "3px 4px", color: theme.textPrimary }}>{d.slotLabel}</td>
                            <td style={{
                                padding: "3px 4px",
                                color: typeColors[d.bindingType] ?? theme.textMuted,
                                fontWeight: 500,
                            }}>
                                {typeLabels[d.bindingType] ?? d.bindingType}
                            </td>
                            <td style={{ padding: "3px 4px", fontFamily: "monospace", fontSize: 10, color: theme.textSecondary }}>
                                {d.rawExpression}
                                {d.warning && (
                                    <div style={{ color: "#d83b01", fontFamily: theme.fontFamily, marginTop: 1 }}>
                                        {d.warning}
                                    </div>
                                )}
                            </td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};

// ════════════════════════════════════════════════════════════════════
// Helper: convert fetch results to SlotField records
// ════════════════════════════════════════════════════════════════════

function toSlotFields(
    results: Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string; color?: string }>,
): Record<string, SlotField> {
    const fields: Record<string, SlotField> = {};
    for (const [key, val] of Object.entries(results)) {
        if (key.startsWith("__")) continue;
        fields[key] = {
            label: val.label,
            value: val.value,
            rawValue: val.value,
            isEmpty: !val.value || val.value === "---",
            lookupEntityType: val.lookupEntityType,
            lookupId: val.lookupId,
            optionColor: val.color,
        };
    }
    return fields;
}

// ════════════════════════════════════════════════════════════════════
// Main Component
// ════════════════════════════════════════════════════════════════════

export const InfoCardComponent: React.FC<InfoCardProps> = (props) => {
    const {
        data,
        layout,
        hideEmpty,
        showBorder,
        showVersion,
        theme,
        version,
        relatedMappings,
        currentRecordMappings,
        currentRecordEntityType,
        currentRecordId,
        fetchRelatedData,
        onOpenRecord,
        bindingDiagnostics,
    } = props;

    // Merge default English strings with any overrides supplied by the host.
    // index.ts builds props.strings from context.resources.getString() so the
    // user's Dataverse UI language wins when the resx is registered.
    const strings: InfoCardStrings = React.useMemo(
        () => ({ ...DEFAULT_STRINGS, ...(props.strings ?? {}) }),
        [props.strings],
    );

    const [relatedFields, setRelatedFields] = React.useState<Record<string, SlotField>>({});
    const [currentRecordFields, setCurrentRecordFields] = React.useState<Record<string, SlotField>>({});
    // Overrides for bound field labels, values, and colors (from record fetch + metadata)
    const [recordOverrides, setRecordOverrides] = React.useState<Record<string, { label: string; value: string; color?: string }>>({});
    // Loading state — true while related-record fetches are in flight
    const [titleFetchDone, setTitleFetchDone] = React.useState(false);
    const [currentFetchDone, setCurrentFetchDone] = React.useState(false);
    // Smart-layout collapse state lives here so the whole card surface can toggle.
    const [collapsed, setCollapsed] = React.useState(!props.startExpanded);
    const toggleCollapsed = React.useCallback(() => setCollapsed(c => !c), []);

    // Resolve bound field labels/values/colors via record fetch
    React.useEffect(() => {
        if (!props.resolveRecordFields) return;
        props.resolveRecordFields().then(setRecordOverrides).catch(() => { /* optional */ });
    }, [props.resolveRecordFields]);

    // Fetch title-entity related data
    React.useEffect(() => {
        if (relatedMappings.length === 0 || !fetchRelatedData) {
            setTitleFetchDone(true);
            return;
        }
        setTitleFetchDone(false);

        const sourceField = data.title;
        if (!sourceField || !sourceField.lookupEntityType || !sourceField.lookupId) {
            if (relatedMappings.length > 0 && sourceField) {
                console.warn("[InfoCard] Title mappings exist but title has no lookup data.");
            }
            setTitleFetchDone(true);
            return;
        }

        const columns = relatedMappings.map(m => m.fetchField);
        fetchRelatedData(sourceField.lookupEntityType, sourceField.lookupId, columns)
            .then((results) => {
                setRelatedFields(toSlotFields(results));
            })
            .catch((err) => {
                console.error("[InfoCard] Title-entity fetch failed:", err);
            })
            .finally(() => setTitleFetchDone(true));
    }, [
        data.title?.lookupId,
        relatedMappings.length,
        relatedMappings.map(m => m.fetchField).join(","),
    ]);

    // Fetch current-record related data (@. syntax)
    React.useEffect(() => {
        if (!currentRecordMappings || currentRecordMappings.length === 0 || !fetchRelatedData) {
            setCurrentFetchDone(true);
            return;
        }
        setCurrentFetchDone(false);

        if (!currentRecordEntityType || !currentRecordId) {
            console.warn("[InfoCard] @. mappings exist but no current record context available.");
            setCurrentFetchDone(true);
            return;
        }

        const columns = currentRecordMappings.map(m => m.fetchField);
        fetchRelatedData(currentRecordEntityType, currentRecordId, columns)
            .then((results) => {
                setCurrentRecordFields(toSlotFields(results));
            })
            .catch((err) => {
                console.error("[InfoCard] Current-record fetch failed:", err);
            })
            .finally(() => setCurrentFetchDone(true));
    }, [
        currentRecordEntityType,
        currentRecordId,
        currentRecordMappings?.length,
        currentRecordMappings?.map(m => m.fetchField).join(","),
    ]);

    // Merge all related fields into display data
    const allMappings = [...relatedMappings, ...(currentRecordMappings ?? [])];
    const allRelatedFields = { ...relatedFields, ...currentRecordFields };
    let displayData = allMappings.length > 0 && Object.keys(allRelatedFields).length > 0
        ? mergeRelatedFields(data, allRelatedFields, allMappings)
        : data;

    // Apply record-fetch overrides (labels, formatted values, colors) for bound fields
    // and for SLOT_PRESETS-populated slots. Covers every group so preset placeholders
    // get filled in once the form record fetch resolves.
    if (Object.keys(recordOverrides).length > 0) {
        const applyOverride = (f: SlotField): SlotField => {
            const ov = f.slotKey ? recordOverrides[f.slotKey] : undefined;
            if (!ov) return f;
            return {
                ...f,
                label: ov.label || f.label,
                ...(ov.value ? { value: ov.value, isEmpty: false } : {}),
                optionColor: ov.color ?? f.optionColor,
            };
        };
        const applyOverrideOrNull = (f: SlotField | null): SlotField | null =>
            f ? applyOverride(f) : f;
        displayData = {
            ...displayData,
            subtitles: displayData.subtitles.map(applyOverride),
            phones: displayData.phones.map(applyOverride),
            email: applyOverrideOrNull(displayData.email),
            web: applyOverrideOrNull(displayData.web),
            address: applyOverrideOrNull(displayData.address),
            details: displayData.details.map(applyOverride),
            gridFields: displayData.gridFields.map((f) => {
                const ov = f.slotKey ? recordOverrides[f.slotKey] : undefined;
                if (!ov) return f;
                return {
                    ...f,
                    label: ov.label || f.label,
                    ...(ov.value ? { value: ov.value, isEmpty: false } : {}),
                    optionColor: ov.color ?? f.optionColor,
                };
            }),
            tags: displayData.tags.map((f) => {
                const ov = f.slotKey ? recordOverrides[f.slotKey] : undefined;
                if (!ov) return f;
                return {
                    ...f,
                    ...(ov.label ? { label: ov.label } : {}),
                    ...(ov.value ? { value: ov.value, isEmpty: false } : {}),
                    optionColor: ov.color ?? f.optionColor,
                };
            }),
        };
        const imageOv = recordOverrides["imageField"];
        if (imageOv && imageOv.value && !displayData.imageUrl) {
            displayData = { ...displayData, imageUrl: imageOv.value };
        }
    }

    // Hide preset-populated slots that remain empty after the override pass — they
    // were speculative fills for columns that have no data on this specific record.
    {
        const dropEmptyPreset = (f: SlotField): boolean => !(f.isPreset && f.isEmpty);
        const dropEmptyPresetOrKeep = (f: SlotField | null): SlotField | null =>
            f && f.isPreset && f.isEmpty ? null : f;
        displayData = {
            ...displayData,
            subtitles: displayData.subtitles.filter(dropEmptyPreset),
            phones: displayData.phones.filter(dropEmptyPreset),
            email: dropEmptyPresetOrKeep(displayData.email),
            web: dropEmptyPresetOrKeep(displayData.web),
            address: dropEmptyPresetOrKeep(displayData.address),
            details: displayData.details.filter(dropEmptyPreset),
            gridFields: displayData.gridFields.filter(dropEmptyPreset),
            tags: displayData.tags.filter(dropEmptyPreset),
        };
    }

    // Compute set of slot names whose related-record fetch is still in flight.
    // Pending slots that are still empty get rendered as a shimmer placeholder.
    const pendingTargets = new Set<string>();
    if (fetchRelatedData && !titleFetchDone) {
        for (const m of relatedMappings) pendingTargets.add(m.targetSlot);
    }
    if (fetchRelatedData && !currentFetchDone) {
        for (const m of (currentRecordMappings ?? [])) pendingTargets.add(m.targetSlot);
    }
    const isLoading = pendingTargets.size > 0;

    if (isLoading) {
        const markPending = (f: SlotField | null, slotKey: string): SlotField | null => {
            if (!f) return f;
            if (pendingTargets.has(slotKey) && f.isEmpty) {
                return { ...f, isPending: true };
            }
            return f;
        };
        displayData = {
            ...displayData,
            subtitles: displayData.subtitles.map((f, i) => markPending(f, `subtitleField${i + 1}`)!),
            phones: displayData.phones.map((f, i) => markPending(f, `phoneField${i + 1}`)!),
            email: markPending(displayData.email, "emailField"),
            web: markPending(displayData.web, "webField"),
            address: markPending(displayData.address, "addressField"),
            details: displayData.details.map((f, i) => markPending(f, `detailField${i + 1}`)!),
            gridFields: displayData.gridFields.map((f, i) => markPending(f, `gridField${i + 1}`)!),
            tags: displayData.tags.map((f, i) => markPending(f, `tagField${i + 1}`)!),
        };
    }

    // Check if we have a valid title to display
    const hasTitle = props.designTime
        ? displayData.title != null
        : displayData.title && !displayData.title.isEmpty;

    // Card wrapper styles
    const cardStyle: React.CSSProperties = {
        width: "100%",
        boxSizing: "border-box" as const,
        overflow: "hidden",
        wordBreak: "break-word" as const,
        background: theme.cardBg,
        fontFamily: theme.fontFamily,
        padding: "12px 14px",
        borderRadius: theme.radius,
        position: "relative",
        ...(showBorder
            ? {
                border: `1px solid ${theme.border}`,
                boxShadow: theme.shadow,
            }
            : {
                border: "none",
                boxShadow: "none",
            }),
    };

    // No data state — but show diagnostics at design time
    if (!hasTitle) {
        return (
            <div style={cardStyle}>
                <div
                    style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        padding: "16px 0",
                        color: theme.textMuted,
                        fontSize: 13,
                        gap: 6,
                    }}
                >
                    <PersonIcon size={16} color={theme.textMuted} />
                    <span>No fields bound</span>
                </div>
                {bindingDiagnostics && bindingDiagnostics.length > 0 && (
                    <DesignTimeBindingPanel diagnostics={bindingDiagnostics} theme={theme} />
                )}
                {showVersion && <VersionBadge version={version} theme={theme} />}
            </div>
        );
    }

    // Determine which sections collapse based on the maker's setting.
    // Each section needs populated content for the collapse to be meaningful.
    const collapseSetting = props.collapsibleSections ?? "body";
    const populatedSections = {
        contact: displayData.phones.some(f => !f.isEmpty || f.isPending) ||
            (!!displayData.email && (!displayData.email.isEmpty || displayData.email.isPending)) ||
            (!!displayData.web && (!displayData.web.isEmpty || displayData.web.isPending)) ||
            (!!displayData.address && (!displayData.address.isEmpty || displayData.address.isPending)),
        body: displayData.details.some(f => !f.isEmpty || f.isPending) ||
            displayData.gridFields.some(f => !f.isEmpty || f.isPending),
        tags: displayData.tags.some(f => !f.isEmpty || f.isPending),
    };
    const isCardCollapsible =
        collapseSetting !== "none" &&
        ((shouldCollapseSection(collapseSetting, "contact") && populatedSections.contact) ||
            (shouldCollapseSection(collapseSetting, "body") && populatedSections.body) ||
            (shouldCollapseSection(collapseSetting, "tags") && populatedSections.tags));

    // Render the appropriate layout
    const layoutProps: LayoutProps = {
        data: displayData,
        theme,
        hideEmpty,
        showTitle: props.showTitle,
        startExpanded: props.startExpanded,
        designTime: props.designTime,
        onOpenRecord,
        collapsed,
        isCollapsible: isCardCollapsible,
        collapsibleSections: collapseSetting,
        strings,
        subtitleSeparator: props.subtitleSeparator,
        titlePrefix: props.titlePrefix,
        imageShape: props.imageShape,
        showDetailIcons: props.showDetailIcons,
        detailLabelStyle: props.detailLabelStyle,
        formFactor: props.formFactor,
    };

    // Whole-card click-to-toggle. Active hit zones (anchors, lookup nav handlers)
    // call stopPropagation so they don't toggle the card.
    const interactiveCardProps: React.HTMLAttributes<HTMLDivElement> = isCardCollapsible ? {
        onClick: toggleCollapsed,
        onKeyDown: (e) => {
            if (e.target !== e.currentTarget) return;
            if (e.key === "Enter" || e.key === " ") {
                e.preventDefault();
                toggleCollapsed();
            }
        },
        role: "button",
        tabIndex: 0,
        "aria-expanded": !collapsed,
        "aria-label": collapsed ? strings.cardExpand : strings.cardCollapse,
    } : {};

    const finalCardStyle: React.CSSProperties = isCardCollapsible
        ? { ...cardStyle, cursor: "pointer", outline: "none" }
        : cardStyle;

    return (
        <div style={finalCardStyle} {...interactiveCardProps}>
            <LoadingKeyframes />
            {isLoading && <TopProgressBar theme={theme} />}
            {layout === "smart" && <SmartCardLayout {...layoutProps} />}
            {layout === "contact" && <ContactCardLayout {...layoutProps} />}
            {layout === "compact" && <CompactCardLayout {...layoutProps} />}
            {bindingDiagnostics && bindingDiagnostics.length > 0 && (
                <DesignTimeBindingPanel diagnostics={bindingDiagnostics} theme={theme} />
            )}
            {showVersion && <VersionBadge version={version} theme={theme} />}
        </div>
    );
};
