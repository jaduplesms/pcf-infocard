/**
 * InfoCard — A compact info card React component for PCF virtual controls.
 *
 * Renders slot-based fields in three layout modes: smart (collapsible),
 * contact (full card), and compact (dense grid). All styling is inline,
 * all icons are inline SVGs. No external dependencies.
 */
import * as React from "react";

// ════════════════════════════════════════════════════════════════════
// Types
// ════════════════════════════════════════════════════════════════════

export type LayoutMode = "smart" | "contact" | "compact";

export interface SlotField {
    label: string;
    value: string;
    rawValue: unknown;
    isEmpty: boolean;
    lookupEntityType?: string;
    lookupId?: string;
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

export interface InfoCardProps {
    data: InfoCardData;
    layout: LayoutMode;
    hideEmpty: boolean;
    showBorder: boolean;
    showVersion: boolean;
    startExpanded: boolean;
    theme: InfoCardTheme;
    version: string;
    relatedMappings: RelatedFieldMapping[];
    fetchRelatedData?: (entityType: string, id: string, columns: string[]) => Promise<Record<string, { value: string; label: string; lookupId?: string; lookupEntityType?: string }>>;
    onOpenRecord?: (entityType: string, id: string) => void;
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
        <path d="M8 1a7 7 0 100 14A7 7 0 008 1zM6.5 2.3C5.6 3.5 5 5.1 4.8 7H2.1a6 6 0 014.4-4.7zM2.1 9h2.7c0.2 1.9 0.8 3.5 1.7 4.7A6 6 0 012.1 9zm4.7 4.7c-1.1-1.1-1.8-2.8-2-4.7h6.4c-0.2 1.9-0.9 3.6-2 4.7-0.4 0.2-0.8 0.3-1.2 0.3s-0.8-0.1-1.2-0.3zM11.2 7c-0.2-1.9-0.9-3.5-2-4.7 0.4-0.2 0.8-0.3 1.2-0.3 2.4 0 4.4 1.5 5.2 3.6A6 6 0 0013.9 7h-2.7zm-6.4 0c0.2-1.9 0.9-3.6 2-4.7C7.2 2.1 7.6 2 8 2s0.8 0.1 1.2 0.3c1.1 1.1 1.8 2.8 2 4.7H4.8zm4.7 2h2.6a6 6 0 01-4.4 4.7c1-1.2 1.6-2.8 1.8-4.7z" fill={color} />
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
    <svg width={size} height={size} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M8 2a3 3 0 100 6 3 3 0 000-6zM3 12c0-2.2 2.2-4 5-4s5 1.8 5 4v1H3v-1z" fill={color} />
    </svg>
);

const ChevronDown: React.FC<IconProps> = ({ size = 12, color = "currentColor" }) => (
    <svg width={size} height={size} viewBox="0 0 12 12" fill="none" xmlns="http://www.w3.org/2000/svg">
        <path d="M2.5 4.5L6 8l3.5-3.5" stroke={color} strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
);

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
    return fields.filter(f => !f.isEmpty);
}

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

function deduplicateSubtitles(subtitles: SlotField[], mutedColor: string): SlotField[] {
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

function mergeRelatedFields(
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

    for (const mapping of mappings) {
        const fetched = relatedFields[mapping.fetchField];
        if (!fetched) continue;

        const target = mapping.targetSlot;

        if (target === "addressField") {
            merged.address = fetched;
        } else if (target === "phoneField1") {
            if (merged.phones.length >= 1) {
                merged.phones[0] = fetched;
            } else {
                merged.phones.push(fetched);
            }
        } else if (target === "phoneField2") {
            while (merged.phones.length < 2) {
                merged.phones.push({ label: "", value: "---", rawValue: null, isEmpty: true });
            }
            merged.phones[1] = fetched;
        } else if (target === "emailField") {
            merged.email = fetched;
        } else if (target === "webField") {
            merged.web = fetched;
        } else if (target === "image") {
            merged.imageUrl = fetched.value;
        } else if (target.startsWith("subtitleField")) {
            const idx = parseInt(target.replace("subtitleField", ""), 10) - 1;
            while (merged.subtitles.length <= idx) {
                merged.subtitles.push({ label: "", value: "---", rawValue: null, isEmpty: true });
            }
            merged.subtitles[idx] = fetched;
        } else if (target.startsWith("detailField")) {
            const idx = parseInt(target.replace("detailField", ""), 10) - 1;
            while (merged.details.length <= idx) {
                merged.details.push({ label: "", value: "---", rawValue: null, isEmpty: true });
            }
            merged.details[idx] = fetched;
        } else if (target.startsWith("gridField")) {
            const idx = parseInt(target.replace("gridField", ""), 10) - 1;
            while (merged.gridFields.length <= idx) {
                merged.gridFields.push({ label: "", value: "---", rawValue: null, isEmpty: true });
            }
            merged.gridFields[idx] = fetched;
        } else if (target.startsWith("tagField")) {
            const idx = parseInt(target.replace("tagField", ""), 10) - 1;
            while (merged.tags.length <= idx) {
                merged.tags.push({ label: "", value: "---", rawValue: null, isEmpty: true });
            }
            merged.tags[idx] = fetched;
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
    onOpenRecord?: (entityType: string, id: string) => void;
}

const Header: React.FC<HeaderProps> = ({ data, theme, hideEmpty, onOpenRecord }) => {
    const title = data.title;
    if (!title || title.isEmpty) return null;

    const isLookup = !!(title.lookupEntityType && title.lookupId);
    const subtitles = deduplicateSubtitles(
        filterEmpty(data.subtitles, hideEmpty),
        theme.textMuted,
    );

    return (
        <div style={{ marginBottom: 8 }}>
            <div
                style={{
                    fontSize: 16,
                    fontWeight: 600,
                    color: isLookup ? theme.brand : theme.textPrimary,
                    cursor: isLookup ? "pointer" : "default",
                    lineHeight: "22px",
                }}
                onClick={isLookup && onOpenRecord ? () => onOpenRecord(title.lookupEntityType!, title.lookupId!) : undefined}
                title={title.label}
            >
                {title.value}
            </div>
            {subtitles.length > 0 && (
                <div style={{ fontSize: 13, color: theme.textSecondary, lineHeight: "18px", marginTop: 2 }}>
                    {subtitles.map((sub, i) => {
                        const isSubLookup = !!(sub.lookupEntityType && sub.lookupId);
                        return (
                            <React.Fragment key={i}>
                                {i > 0 && (
                                    <span style={{ margin: "0 6px", color: theme.textMuted }}>{"\u00b7"}</span>
                                )}
                                <span
                                    style={{
                                        color: isSubLookup ? theme.brand : theme.textSecondary,
                                        cursor: isSubLookup ? "pointer" : "default",
                                        textDecoration: isSubLookup ? "none" : "none",
                                    }}
                                    onClick={
                                        isSubLookup && onOpenRecord
                                            ? () => onOpenRecord(sub.lookupEntityType!, sub.lookupId!)
                                            : undefined
                                    }
                                    title={sub.label}
                                >
                                    {sub.value}
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
}

const ContactRows: React.FC<ContactRowsProps> = ({ data, theme, hideEmpty }) => {
    const mapUrl = buildMapUrl(data.latitude, data.longitude);
    const address = data.address;
    const phones = filterEmpty(data.phones, hideEmpty);
    const email = data.email && !data.email.isEmpty ? data.email : null;
    const web = data.web && !data.web.isEmpty ? data.web : null;

    const hasAny = (address && !address.isEmpty) || phones.length > 0 || email || web;
    if (!hasAny) return null;

    const rowStyle: React.CSSProperties = {
        display: "flex",
        alignItems: "flex-start",
        gap: 8,
        padding: "4px 0",
        fontSize: 13,
        color: theme.textSecondary,
        lineHeight: "18px",
    };

    const iconWrapStyle: React.CSSProperties = {
        flexShrink: 0,
        marginTop: 2,
    };

    const linkStyle: React.CSSProperties = {
        color: theme.brand,
        textDecoration: "none",
        cursor: "pointer",
    };

    const chipStyle: React.CSSProperties = {
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        padding: "6px 10px",
        fontSize: 13,
        color: theme.textPrimary,
        textDecoration: "none",
        whiteSpace: "nowrap",
        borderRadius: 6,
        cursor: "pointer",
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
            {address && !address.isEmpty && (
                mapUrl ? (
                    <a href={mapUrl} target="_blank" rel="noopener noreferrer" style={chipStyle} title={address.label}>
                        <PinIcon size={14} color={theme.brand} />
                        {address.value}
                    </a>
                ) : (
                    <span style={{ ...chipStyle, color: theme.textSecondary, cursor: "default" }} title={address.label}>
                        <PinIcon size={14} color={theme.textMuted} />
                        {address.value}
                    </span>
                )
            )}

            {/* Phones — phone1 gets landline icon, phone2 gets mobile icon */}
            {phones.map((phone, i) => (
                <a key={`phone-${i}`} href={`tel:${String(phone.value).replace(/\s+/g, "")}`} style={chipStyle} title={phone.label}>
                    {i === 0 ? <PhoneIcon size={14} color={theme.brand} /> : <MobileIcon size={14} color={theme.brand} />}
                    {phone.value}
                </a>
            ))}

            {/* Email */}
            {email && (
                <a href={`mailto:${email.value}`} style={chipStyle} title={email.label}>
                    <EmailIcon size={14} color={theme.brand} />
                    {email.value}
                </a>
            )}

            {/* Web */}
            {web && (
                <a
                    href={web.value.startsWith("http") ? web.value : `https://${web.value}`}
                    target="_blank"
                    rel="noopener noreferrer"
                    style={chipStyle}
                    title={web.label}
                >
                    <WebIcon size={14} color={theme.brand} />
                    {web.value}
                </a>
            )}
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
}

const DetailRows: React.FC<DetailRowsProps> = ({ details, theme, hideEmpty, latitude, longitude }) => {
    const filtered = filterEmpty(details, hideEmpty);
    if (filtered.length === 0) return null;

    const mapUrl = buildMapUrl(latitude, longitude);

    return (
        <div style={{ display: "flex", flexDirection: "column", gap: 0 }}>
            {filtered.map((field, i) => {
                const icon = guessDetailIcon(field, theme.textMuted);
                const isAddressLike = field.label.toLowerCase().includes("address") || field.label.toLowerCase().includes("location");
                const showMapLink = isAddressLike && mapUrl;

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
                            <a
                                href={mapUrl!}
                                target="_blank"
                                rel="noopener noreferrer"
                                style={{ color: theme.brand, textDecoration: "none", cursor: "pointer" }}
                            >
                                {field.value}
                            </a>
                        ) : (
                            <span style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>{field.value}</span>
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
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                        }}
                        title={field.value}
                    >
                        {field.value}
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
            {filtered.map((tag, i) => (
                <span
                    key={i}
                    style={{
                        display: "inline-block",
                        padding: "2px 8px",
                        fontSize: 11,
                        fontWeight: 500,
                        color: theme.brand,
                        background: theme.brandLight,
                        borderRadius: "10px",
                        lineHeight: "16px",
                        whiteSpace: "nowrap",
                    }}
                    title={tag.label}
                >
                    {tag.value}
                </span>
            ))}
        </div>
    );
};

// ── Shared: Image ─────────────────────────────────────────────────

interface ImageProps {
    imageUrl: string | null;
    title: string;
}

const ImageAvatar: React.FC<ImageProps> = ({ imageUrl, title }) => {
    if (!imageUrl) return null;
    return (
        <img
            src={imageUrl}
            alt={title}
            style={{
                width: 48,
                height: 48,
                borderRadius: "50%",
                objectFit: "cover",
                flexShrink: 0,
            }}
        />
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
    startExpanded?: boolean;
    onOpenRecord?: (entityType: string, id: string) => void;
}

const SmartCardLayout: React.FC<LayoutProps> = ({ data, theme, hideEmpty, startExpanded = true, onOpenRecord }) => {
    const [collapsed, setCollapsed] = React.useState(!startExpanded);

    const details = filterEmpty(data.details, hideEmpty);
    const gridFields = filterEmpty(data.gridFields, hideEmpty);
    const tags = filterEmpty(data.tags, hideEmpty);

    const hasBody = details.length > 0 || gridFields.length > 0;
    const hasCollapsible = hasBody;

    return (
        <div>
            {/* Header row with image, title, subtitles, and chevron */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                <ImageAvatar imageUrl={data.imageUrl} title={data.title?.value ?? ""} />
                <div style={{ flex: 1, minWidth: 0 }}>
                    <Header data={data} theme={theme} hideEmpty={hideEmpty} onOpenRecord={onOpenRecord} />
                </div>
                {hasCollapsible && (
                    <div
                        style={{
                            cursor: "pointer",
                            padding: 4,
                            flexShrink: 0,
                            transition: "transform 0.2s ease",
                            transform: collapsed ? "rotate(-90deg)" : "rotate(0deg)",
                        }}
                        onClick={() => setCollapsed(!collapsed)}
                    >
                        <ChevronDown size={12} color={theme.textMuted} />
                    </div>
                )}
            </div>

            {/* Contact section — ALWAYS visible */}
            <ContactRows data={data} theme={theme} hideEmpty={hideEmpty} />

            {/* Collapsible body */}
            {!collapsed && hasBody && (
                <div style={{ marginTop: 8 }}>
                    {/* Detail rows */}
                    {details.length > 0 && (
                        <div style={{ marginBottom: gridFields.length > 0 ? 8 : 0 }}>
                            <DetailRows
                                details={data.details}
                                theme={theme}
                                hideEmpty={hideEmpty}
                                latitude={data.latitude}
                                longitude={data.longitude}
                            />
                        </div>
                    )}

                    {/* Grid fields */}
                    {gridFields.length > 0 && (
                        <div style={{ marginTop: details.length > 0 ? 0 : 0 }}>
                            <GridFields fields={data.gridFields} theme={theme} hideEmpty={hideEmpty} />
                        </div>
                    )}
                </div>
            )}

            {/* Tags — ALWAYS visible */}
            {tags.length > 0 && (
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

const ContactCardLayout: React.FC<LayoutProps> = ({ data, theme, hideEmpty, onOpenRecord }) => {
    const details = filterEmpty(data.details, hideEmpty);
    const gridFields = filterEmpty(data.gridFields, hideEmpty);
    const tags = filterEmpty(data.tags, hideEmpty);

    return (
        <div>
            {/* Header row with image */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                <ImageAvatar imageUrl={data.imageUrl} title={data.title?.value ?? ""} />
                <div style={{ flex: 1, minWidth: 0 }}>
                    <Header data={data} theme={theme} hideEmpty={hideEmpty} onOpenRecord={onOpenRecord} />
                </div>
            </div>

            {/* Contact section */}
            <ContactRows data={data} theme={theme} hideEmpty={hideEmpty} />

            {/* Detail rows */}
            {details.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    <DetailRows
                        details={data.details}
                        theme={theme}
                        hideEmpty={hideEmpty}
                        latitude={data.latitude}
                        longitude={data.longitude}
                    />
                </div>
            )}

            {/* Grid fields */}
            {gridFields.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    <GridFields fields={data.gridFields} theme={theme} hideEmpty={hideEmpty} />
                </div>
            )}

            {/* Tags */}
            {tags.length > 0 && (
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

const CompactCardLayout: React.FC<LayoutProps> = ({ data, theme, hideEmpty, onOpenRecord }) => {
    const phones = filterEmpty(data.phones, hideEmpty);
    const email = data.email && !data.email.isEmpty ? data.email : null;
    const web = data.web && !data.web.isEmpty ? data.web : null;
    const address = data.address && !data.address.isEmpty ? data.address : null;
    const details = filterEmpty(data.details, hideEmpty);
    const gridFields = filterEmpty(data.gridFields, hideEmpty);
    const tags = filterEmpty(data.tags, hideEmpty);

    const hasContact = phones.length > 0 || email || web || address;
    const hasDetails = gridFields.length > 0;
    const hasInfo = details.length > 0;

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
            {/* Header */}
            <div style={{ display: "flex", alignItems: "flex-start", gap: 10 }}>
                <ImageAvatar imageUrl={data.imageUrl} title={data.title?.value ?? ""} />
                <div style={{ flex: 1, minWidth: 0 }}>
                    <Header data={data} theme={theme} hideEmpty={hideEmpty} onOpenRecord={onOpenRecord} />
                </div>
            </div>

            {/* Contact section */}
            {hasContact && (
                <div>
                    <div style={sectionHeaderStyle}>Contact</div>
                    {address && (
                        <div style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{address.label}</span>
                            <span style={{ color: theme.textPrimary, textAlign: "right" }}>{address.value}</span>
                        </div>
                    )}
                    {phones.map((phone, i) => (
                        <div key={`phone-${i}`} style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{phone.label}</span>
                            <a href={`tel:${String(phone.value).replace(/\s+/g, "")}`} style={{ color: theme.brand, textDecoration: "none" }}>
                                {phone.value}
                            </a>
                        </div>
                    ))}
                    {email && (
                        <div style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{email.label}</span>
                            <a href={`mailto:${email.value}`} style={{ color: theme.brand, textDecoration: "none" }}>
                                {email.value}
                            </a>
                        </div>
                    )}
                    {web && (
                        <div style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{web.label}</span>
                            <a
                                href={web.value.startsWith("http") ? web.value : `https://${web.value}`}
                                target="_blank"
                                rel="noopener noreferrer"
                                style={{ color: theme.brand, textDecoration: "none" }}
                            >
                                {web.value}
                            </a>
                        </div>
                    )}
                </div>
            )}

            {/* Details section (grid fields) */}
            {hasDetails && (
                <div>
                    <div style={sectionHeaderStyle}>Details</div>
                    {gridFields.map((field, i) => (
                        <div key={i} style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{field.label}</span>
                            <span style={{ color: theme.textPrimary, textAlign: "right" }}>{field.value}</span>
                        </div>
                    ))}
                </div>
            )}

            {/* Info section (detail fields) */}
            {hasInfo && (
                <div>
                    <div style={sectionHeaderStyle}>Info</div>
                    {details.map((field, i) => (
                        <div key={i} style={fieldRowStyle}>
                            <span style={{ color: theme.textMuted }}>{field.label}</span>
                            <span style={{ color: theme.textPrimary, textAlign: "right", wordBreak: "break-word" }}>
                                {field.value}
                            </span>
                        </div>
                    ))}
                </div>
            )}

            {/* Tags */}
            {tags.length > 0 && (
                <div style={{ marginTop: 8 }}>
                    <Tags tags={data.tags} theme={theme} hideEmpty={hideEmpty} />
                </div>
            )}
        </div>
    );
};

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
        fetchRelatedData,
        onOpenRecord,
    } = props;

    const [relatedFields, setRelatedFields] = React.useState<Record<string, SlotField>>({});

    // Fetch related data when mappings or source lookup change
    React.useEffect(() => {
        if (relatedMappings.length === 0 || !fetchRelatedData) return;

        const sourceField = data.title;
        if (!sourceField || !sourceField.lookupEntityType || !sourceField.lookupId) return;

        const columns = relatedMappings.map(m => m.fetchField);
        const entityType = sourceField.lookupEntityType;
        const entityId = sourceField.lookupId;

        fetchRelatedData(entityType, entityId, columns)
            .then((results) => {
                // Convert to SlotField records
                const fields: Record<string, SlotField> = {};
                for (const [key, val] of Object.entries(results)) {
                    if (key.startsWith("__")) continue; // skip debug keys
                    fields[key] = {
                        label: val.label,
                        value: val.value,
                        rawValue: val.value,
                        isEmpty: !val.value || val.value === "---",
                        lookupEntityType: val.lookupEntityType,
                        lookupId: val.lookupId,
                    };
                }
                setRelatedFields(fields);
            })
            .catch(() => {
                // Silently handle errors
            });
    }, [
        data.title?.lookupId,
        relatedMappings.length,
        relatedMappings.map(m => m.fetchField).join(","),
    ]);

    // Merge related fields into display data
    const displayData = relatedMappings.length > 0 && Object.keys(relatedFields).length > 0
        ? mergeRelatedFields(data, relatedFields, relatedMappings)
        : data;

    // Check if we have a valid title to display
    const hasTitle = displayData.title && !displayData.title.isEmpty;

    // Card wrapper styles
    const cardStyle: React.CSSProperties = {
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

    // No data state
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
                {showVersion && <VersionBadge version={version} theme={theme} />}
            </div>
        );
    }

    // Render the appropriate layout
    const layoutProps: LayoutProps = {
        data: displayData,
        theme,
        hideEmpty,
        startExpanded: props.startExpanded,
        onOpenRecord,
    };

    return (
        <div style={cardStyle}>
            {layout === "smart" && <SmartCardLayout {...layoutProps} />}
            {layout === "contact" && <ContactCardLayout {...layoutProps} />}
            {layout === "compact" && <CompactCardLayout {...layoutProps} />}
            {showVersion && <VersionBadge version={version} theme={theme} />}
        </div>
    );
};
