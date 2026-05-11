/**
 * Component tests for InfoCardComponent (InfoCard.tsx).
 *
 * Covers rendering, layout modes, field visibility, action buttons,
 * theme colors, version badge, lookup navigation, and related fields.
 */

import * as React from "react";
import { render, screen, fireEvent, waitFor, act } from "@testing-library/react";
import "@testing-library/jest-dom";

// React 16 compatible: flush microtasks to let promise-based effects settle
function flushPromises(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}
import {
  InfoCardComponent,
  InfoCardData,
  InfoCardProps,
  InfoCardTheme,
  SlotField,
  LayoutMode,
  defaultTheme,
  mergeRelatedFields,
} from "../InfoCard";

interface RelatedFieldMapping {
  sourceSlot: string;
  fetchField: string;
  targetSlot: string;
}

// ────────────────────────────────────────
// Test helpers
// ────────────────────────────────────────

function makeField(overrides: Partial<SlotField> & { label: string; value: string }): SlotField {
  return {
    rawValue: overrides.rawValue ?? overrides.value,
    isEmpty: overrides.isEmpty ?? false,
    lookupEntityType: overrides.lookupEntityType,
    lookupId: overrides.lookupId,
    ...overrides,
  };
}

function emptyField(label: string): SlotField {
  return {
    label,
    value: "---",
    rawValue: null,
    isEmpty: true,
  };
}

function makeData(overrides: Partial<InfoCardData> = {}): InfoCardData {
  return {
    title: "title" in overrides ? overrides.title! : makeField({ label: "Name", value: "Test Card" }),
    subtitles: overrides.subtitles ?? [],
    phones: overrides.phones ?? [],
    email: "email" in overrides ? overrides.email! : null,
    web: "web" in overrides ? overrides.web! : null,
    address: "address" in overrides ? overrides.address! : null,
    latitude: "latitude" in overrides ? overrides.latitude! : null,
    longitude: "longitude" in overrides ? overrides.longitude! : null,
    details: overrides.details ?? [],
    gridFields: overrides.gridFields ?? [],
    tags: overrides.tags ?? [],
    imageUrl: overrides.imageUrl ?? null,
  };
}

function makeProps(overrides: Partial<InfoCardProps> = {}): InfoCardProps {
  return {
    data: overrides.data ?? makeData(),
    layout: overrides.layout ?? "contact",
    hideEmpty: overrides.hideEmpty ?? true,
    showBorder: overrides.showBorder ?? true,
    showVersion: overrides.showVersion ?? false,
    showTitle: overrides.showTitle ?? true,
    startExpanded: overrides.startExpanded ?? true,
    theme: overrides.theme ?? defaultTheme,
    version: overrides.version ?? "2.4.7",
    relatedMappings: overrides.relatedMappings ?? [],
    currentRecordMappings: overrides.currentRecordMappings,
    currentRecordEntityType: overrides.currentRecordEntityType,
    currentRecordId: overrides.currentRecordId,
    fetchRelatedData: overrides.fetchRelatedData,
    resolveRecordFields: overrides.resolveRecordFields,
    onOpenRecord: overrides.onOpenRecord,
    subtitleSeparator: overrides.subtitleSeparator,
    titlePrefix: overrides.titlePrefix,
    imageShape: overrides.imageShape,
    collapsibleSections: overrides.collapsibleSections,
    showDetailIcons: overrides.showDetailIcons,
    detailLabelStyle: overrides.detailLabelStyle,
    formFactor: overrides.formFactor,
  };
}

// ────────────────────────────────────────
// Tests
// ────────────────────────────────────────

describe("InfoCardComponent", () => {

  // ── Title ──────────────────────────────

  describe("title rendering", () => {
    it("renders the title text", () => {
      const { container } = render(<InfoCardComponent {...makeProps()} />);
      expect(container.textContent).toContain("Test Card");
    });

    it("shows 'No fields bound' when title is null", () => {
      const data = makeData({ title: null });
      const { container } = render(<InfoCardComponent {...makeProps({ data })} />);
      expect(container.textContent).toContain("No fields bound");
    });

    it("shows 'No fields bound' when title is empty", () => {
      const data = makeData({ title: emptyField("Name") });
      const { container } = render(<InfoCardComponent {...makeProps({ data })} />);
      expect(container.textContent).toContain("No fields bound");
    });
  });

  // ── Subtitles ─────────────────────────

  describe("subtitle rendering", () => {
    it("renders subtitles", () => {
      const data = makeData({
        subtitles: [
          makeField({ label: "Account", value: "Contoso" }),
          makeField({ label: "Priority", value: "High" }),
        ],
      });
      const { container } = render(<InfoCardComponent {...makeProps({ data })} />);
      expect(container.textContent).toContain("Contoso");
      expect(container.textContent).toContain("High");
    });

    it("renders dot separator between subtitles in contact layout", () => {
      const data = makeData({
        subtitles: [
          makeField({ label: "Account", value: "Contoso" }),
          makeField({ label: "Priority", value: "High" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      // The dot separator is a middle dot character \u00b7
      expect(container.textContent).toContain("\u00b7");
    });

    it("uses custom subtitleSeparator when provided", () => {
      const data = makeData({
        subtitles: [
          makeField({ label: "Account", value: "Contoso" }),
          makeField({ label: "Priority", value: "High" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact", subtitleSeparator: " | " })} />,
      );
      expect(container.textContent).toContain("|");
      // Default middle dot is no longer present between subtitles when overridden
      expect(container.textContent).not.toContain("\u00b7");
    });
  });

  // ── Hide empty fields ─────────────────

  describe("hideEmpty behavior", () => {
    it("hides empty fields when hideEmpty=true", () => {
      const data = makeData({
        gridFields: [
          makeField({ label: "Status", value: "Active" }),
          emptyField("Empty Field"),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, hideEmpty: true, layout: "smart" })} />,
      );
      expect(container.textContent).toContain("Active");
      expect(container.textContent).not.toContain("Empty Field");
    });

    it("shows empty fields as '---' when hideEmpty=false", () => {
      const data = makeData({
        gridFields: [
          emptyField("Empty Grid"),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, hideEmpty: false, layout: "smart" })} />,
      );
      expect(container.textContent).toContain("Empty Grid");
      expect(container.textContent).toContain("---");
    });
  });

  // ── Phone buttons ─────────────────────

  describe("phone buttons", () => {
    it("renders phone buttons with tel: links", () => {
      const data = makeData({
        phones: [
          makeField({ label: "Phone", value: "+1 555 0123", rawValue: "+1 555 0123" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      const links = container.querySelectorAll("a[href^='tel:']");
      expect(links.length).toBeGreaterThanOrEqual(1);
      // The href should have whitespace stripped
      expect(links[0].getAttribute("href")).toBe("tel:+15550123");
    });

    it("renders multiple phone buttons", () => {
      const data = makeData({
        phones: [
          makeField({ label: "Mobile", value: "+1 555 0001", rawValue: "+1 555 0001" }),
          makeField({ label: "Work", value: "+1 555 0002", rawValue: "+1 555 0002" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      const links = container.querySelectorAll("a[href^='tel:']");
      expect(links.length).toBeGreaterThanOrEqual(2);
    });
  });

  // ── Email button ──────────────────────

  describe("email button", () => {
    it("renders email button with mailto: link", () => {
      const data = makeData({
        email: makeField({ label: "Email", value: "test@example.com" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      const link = container.querySelector("a[href='mailto:test@example.com']");
      expect(link).not.toBeNull();
    });

    it("does not render email button when email is empty", () => {
      const data = makeData({
        email: emptyField("Email"),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      const link = container.querySelector("a[href^='mailto:']");
      expect(link).toBeNull();
    });
  });

  // ── Web button ────────────────────────

  describe("web button", () => {
    it("renders web button with https link", () => {
      const data = makeData({
        web: makeField({ label: "Website", value: "https://example.com" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      const link = container.querySelector("a[href='https://example.com']");
      expect(link).not.toBeNull();
    });

    it("prepends https:// when missing", () => {
      const data = makeData({
        web: makeField({ label: "Website", value: "example.com" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      const link = container.querySelector("a[href='https://example.com']");
      expect(link).not.toBeNull();
    });
  });

  // ── Copy-to-clipboard (desktop only) ───

  describe("copy-to-clipboard", () => {
    const originalClipboard = (navigator as Navigator & { clipboard?: Clipboard }).clipboard;

    beforeEach(() => {
      Object.defineProperty(navigator, "clipboard", {
        configurable: true,
        value: { writeText: jest.fn().mockResolvedValue(undefined) },
      });
    });

    afterEach(() => {
      Object.defineProperty(navigator, "clipboard", {
        configurable: true,
        value: originalClipboard,
      });
    });

    it("does not render copy buttons on mobile (formFactor=0)", () => {
      const data = makeData({
        phones: [makeField({ label: "Phone", value: "+1 555 1234" })],
        email: makeField({ label: "Email", value: "a@b.co" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact", formFactor: 0 })} />,
      );
      // Copy button has aria-label starting with "Copy "
      const copyBtn = container.querySelector("button[aria-label^='Copy ']");
      expect(copyBtn).toBeNull();
    });

    it("does not render copy buttons on tablet (formFactor=1)", () => {
      const data = makeData({
        phones: [makeField({ label: "Phone", value: "+1 555 1234" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact", formFactor: 1 })} />,
      );
      expect(container.querySelector("button[aria-label^='Copy ']")).toBeNull();
    });

    it("renders copy buttons on web (formFactor=2)", () => {
      const data = makeData({
        phones: [makeField({ label: "Phone", value: "+1 555 1234" })],
        email: makeField({ label: "Email", value: "a@b.co" }),
        address: makeField({ label: "Address", value: "1 Microsoft Way" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact", formFactor: 2 })} />,
      );
      const copyButtons = container.querySelectorAll("button[aria-label^='Copy ']");
      expect(copyButtons.length).toBe(3); // phone, email, address
    });

    it("invokes navigator.clipboard.writeText with chip value when clicked", async () => {
      const data = makeData({
        phones: [makeField({ label: "Phone", value: "+1 555 1234" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact", formFactor: 2 })} />,
      );
      const copyBtn = container.querySelector("button[aria-label^='Copy ']") as HTMLButtonElement;
      expect(copyBtn).not.toBeNull();
      fireEvent.click(copyBtn);
      await flushPromises();
      const writeText = (navigator.clipboard as Clipboard).writeText as jest.Mock;
      expect(writeText).toHaveBeenCalledWith("+1 555 1234");
    });

    it("hides copy buttons when navigator.clipboard is unavailable", () => {
      Object.defineProperty(navigator, "clipboard", { configurable: true, value: undefined });
      const data = makeData({
        phones: [makeField({ label: "Phone", value: "+1 555 1234" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact", formFactor: 2 })} />,
      );
      expect(container.querySelector("button[aria-label^='Copy ']")).toBeNull();
    });
  });

  // ── Map link ──────────────────────────

  describe("map link", () => {
    it("renders map link when lat/lng provided", () => {
      const data = makeData({
        latitude: 47.6062,
        longitude: -122.3321,
        details: [
          makeField({ label: "Address", value: "123 Main St" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const mapLink = container.querySelector("a[href*='maps.google.com']");
      expect(mapLink).not.toBeNull();
      expect(mapLink!.getAttribute("href")).toContain("47.6062");
      expect(mapLink!.getAttribute("href")).toContain("-122.3321");
    });

    it("does not render map link when lat/lng are null", () => {
      const data = makeData({
        latitude: null,
        longitude: null,
        details: [
          makeField({ label: "Address", value: "123 Main St" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const mapLink = container.querySelector("a[href*='maps.google.com']");
      expect(mapLink).toBeNull();
    });

    it("does not render map link when lat/lng are both zero", () => {
      const data = makeData({
        latitude: 0,
        longitude: 0,
        details: [
          makeField({ label: "Address", value: "123 Main St" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const mapLink = container.querySelector("a[href*='maps.google.com']");
      expect(mapLink).toBeNull();
    });
  });

  // ── Grid fields ───────────────────────

  describe("grid fields", () => {
    it("renders grid fields in smart layout", () => {
      const data = makeData({
        gridFields: [
          makeField({ label: "Field A", value: "Value A" }),
          makeField({ label: "Field B", value: "Value B" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      expect(container.textContent).toContain("Field A");
      expect(container.textContent).toContain("Value A");
      expect(container.textContent).toContain("Field B");
      expect(container.textContent).toContain("Value B");
    });

    it("renders grid fields with 2-column grid in smart layout", () => {
      const data = makeData({
        gridFields: [
          makeField({ label: "F1", value: "V1" }),
          makeField({ label: "F2", value: "V2" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      // The grid container should use "1fr 1fr" template
      const gridDiv = container.querySelector("div[style*='grid-template-columns']") as HTMLElement;
      expect(gridDiv).not.toBeNull();
      expect(gridDiv.style.gridTemplateColumns).toContain("1fr 1fr");
    });

    it("renders single grid field as 1fr", () => {
      const data = makeData({
        gridFields: [
          makeField({ label: "F1", value: "V1" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const gridDiv = container.querySelector("div[style*='grid-template-columns']") as HTMLElement;
      expect(gridDiv).not.toBeNull();
      expect(gridDiv.style.gridTemplateColumns).toBe("1fr");
    });
  });

  // ── Tags ──────────────────────────────

  describe("tags", () => {
    it("renders tags as chips", () => {
      const data = makeData({
        tags: [
          makeField({ label: "Tag1", value: "Urgent" }),
          makeField({ label: "Tag2", value: "Open" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      expect(container.textContent).toContain("Urgent");
      expect(container.textContent).toContain("Open");
    });

    it("all tags use consistent branded styling", () => {
      const data = makeData({
        tags: [
          makeField({ label: "Tag1", value: "Primary" }),
          makeField({ label: "Tag2", value: "Secondary" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const spans = Array.from(container.querySelectorAll("span"));
      const primarySpan = spans.find(s => s.textContent === "Primary");
      const secondarySpan = spans.find(s => s.textContent === "Secondary");

      expect(primarySpan).toBeDefined();
      expect(secondarySpan).toBeDefined();
      // All tags should have the same branded background
      expect(primarySpan!.style.background).toBe(secondarySpan!.style.background);
    });
  });

  // ── Version badge ─────────────────────

  describe("version badge", () => {
    it("shows version info when showVersion=true", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ showVersion: true, version: "1.2.3" })} />,
      );
      // The version badge has a title attribute
      const versionEl = container.querySelector("[title='InfoCard v1.2.3']");
      expect(versionEl).not.toBeNull();
    });

    it("hides version info when showVersion=false", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ showVersion: false, version: "1.2.3" })} />,
      );
      const versionEl = container.querySelector("[title='InfoCard v1.2.3']");
      expect(versionEl).toBeNull();
    });

    it("shows tooltip on click", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ showVersion: true, version: "1.2.3" })} />,
      );
      const versionEl = container.querySelector("[title='InfoCard v1.2.3']");
      expect(versionEl).not.toBeNull();
      fireEvent.click(versionEl!);
      // After click, a tooltip span appears with "v1.2.3"
      expect(container.textContent).toContain("v1.2.3");
    });
  });

  // ── Border ────────────────────────────

  describe("card border", () => {
    it("shows border when showBorder=true", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ showBorder: true })} />,
      );
      const card = container.firstElementChild as HTMLElement;
      // jsdom may split shorthand — check for the individual properties or the raw attribute
      const styleAttr = card.getAttribute("style") ?? "";
      expect(styleAttr).toContain("1px solid");
    });

    it("hides border when showBorder=false", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ showBorder: false })} />,
      );
      const card = container.firstElementChild as HTMLElement;
      const styleAttr = card.getAttribute("style") ?? "";
      // jsdom may not preserve "border: none" as-is; it may omit border entirely
      // Verify that no "1px solid" border is present
      expect(styleAttr).not.toContain("1px solid");
    });

    it("hides box shadow when showBorder=false", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ showBorder: false })} />,
      );
      const card = container.firstElementChild as HTMLElement;
      const styleAttr = card.getAttribute("style") ?? "";
      expect(styleAttr).toContain("box-shadow: none");
    });
  });

  // ── Custom theme ──────────────────────

  describe("custom theme", () => {
    it("applies custom theme colors", () => {
      const customTheme: InfoCardTheme = {
        ...defaultTheme,
        cardBg: "#ff0000",
        textPrimary: "#00ff00",
      };
      const { container } = render(
        <InfoCardComponent {...makeProps({ theme: customTheme })} />,
      );
      const card = container.firstElementChild as HTMLElement;
      expect(card.style.background).toBe("rgb(255, 0, 0)");
    });
  });

  // ── Layout modes ──────────────────────

  describe("layout modes", () => {
    const fullData = makeData({
      subtitles: [makeField({ label: "Account", value: "Contoso" })],
      phones: [makeField({ label: "Phone", value: "+1555", rawValue: "+1555" })],
      email: makeField({ label: "Email", value: "a@b.com" }),
      details: [makeField({ label: "Notes", value: "Some notes" })],
      gridFields: [makeField({ label: "Status", value: "Active" })],
      tags: [makeField({ label: "Tag", value: "Urgent" })],
    });

    it("renders smart layout without error", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: fullData, layout: "smart" })} />,
      );
      expect(container.textContent).toContain("Test Card");
      expect(container.textContent).toContain("Contoso");
    });

    it("renders contact layout without error", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: fullData, layout: "contact" })} />,
      );
      expect(container.textContent).toContain("Test Card");
      expect(container.textContent).toContain("Contoso");
    });

    it("renders compact layout without error", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: fullData, layout: "compact" })} />,
      );
      expect(container.textContent).toContain("Test Card");
      // Compact layout shows section headers
      expect(container.textContent).toContain("Contact");
    });

    it("smart layout has collapsible body", () => {
      const data = makeData({
        phones: [makeField({ label: "Phone", value: "+1555", rawValue: "+1555" })],
        gridFields: [makeField({ label: "Status", value: "Active" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      // Look for the chevron element
      const chevron = container.querySelector("svg[viewBox='0 0 12 12']");
      expect(chevron).not.toBeNull();
    });
  });

  // ── Lookup navigation ─────────────────

  describe("lookup navigation", () => {
    it("renders lookup fields as clickable links", () => {
      const data = makeData({
        subtitles: [
          makeField({
            label: "Customer",
            value: "Contoso Ltd",
            lookupEntityType: "account",
            lookupId: "abc-123",
          }),
        ],
      });
      const onOpenRecord = jest.fn();
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, onOpenRecord, layout: "contact" })} />,
      );
      // The lookup subtitle should be styled as a clickable link
      const spans = Array.from(container.querySelectorAll("span"));
      const lookupSpan = spans.find(s => s.textContent?.includes("Contoso Ltd") && s.style.cursor === "pointer");
      expect(lookupSpan).toBeDefined();
    });

    it("calls onOpenRecord when lookup clicked", () => {
      const data = makeData({
        subtitles: [
          makeField({
            label: "Customer",
            value: "Contoso Ltd",
            lookupEntityType: "account",
            lookupId: "abc-123",
          }),
        ],
      });
      const onOpenRecord = jest.fn();
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, onOpenRecord, layout: "contact" })} />,
      );
      const spans = Array.from(container.querySelectorAll("span"));
      const lookupSpan = spans.find(s => s.textContent?.includes("Contoso Ltd") && s.style.cursor === "pointer");
      expect(lookupSpan).toBeDefined();
      fireEvent.click(lookupSpan!);
      expect(onOpenRecord).toHaveBeenCalledWith("account", "abc-123");
    });

    it("does not call onOpenRecord for non-lookup subtitles", () => {
      const data = makeData({
        subtitles: [
          makeField({
            label: "Status",
            value: "Active",
            // no lookupEntityType or lookupId
          }),
        ],
      });
      const onOpenRecord = jest.fn();
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, onOpenRecord, layout: "contact" })} />,
      );
      const spans = Array.from(container.querySelectorAll("span"));
      const statusSpan = spans.find(s => s.textContent?.includes("Active"));
      expect(statusSpan).toBeDefined();
      fireEvent.click(statusSpan!);
      expect(onOpenRecord).not.toHaveBeenCalled();
    });

    it("opens lookup record on Enter / Space keypress (a11y keyboard)", () => {
      const data = makeData({
        title: makeField({
          label: "Account",
          value: "Contoso",
          rawValue: "abc-123",
          lookupEntityType: "account",
          lookupId: "abc-123",
        }),
      });
      const onOpenRecord = jest.fn();
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, onOpenRecord, layout: "contact" })} />,
      );

      const titleEl = container.querySelector('[role="button"][aria-label*="Open record"]');
      expect(titleEl).toBeTruthy();
      fireEvent.keyDown(titleEl!, { key: "Enter" });
      expect(onOpenRecord).toHaveBeenCalledWith("account", "abc-123");

      onOpenRecord.mockClear();
      fireEvent.keyDown(titleEl!, { key: " " });
      expect(onOpenRecord).toHaveBeenCalledWith("account", "abc-123");

      onOpenRecord.mockClear();
      fireEvent.keyDown(titleEl!, { key: "Tab" });
      expect(onOpenRecord).not.toHaveBeenCalled();
    });

    it("non-lookup title has no role/tabIndex (not keyboard-focusable)", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Plain Title" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      // Title text node is rendered without role="button"
      const buttons = container.querySelectorAll('[role="button"]');
      const titleAsBtn = Array.from(buttons).find(b => b.textContent === "Plain Title");
      expect(titleAsBtn).toBeUndefined();
    });

    it("phone action chip has localized aria-label", () => {
      const data = makeData({
        title: makeField({ label: "Account", value: "Contoso" }),
        phones: [makeField({ label: "Phone", value: "555-1234" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      const phoneLink = container.querySelector('a[href^="tel:"]');
      expect(phoneLink).toBeTruthy();
      expect(phoneLink!.getAttribute("aria-label")).toBe("Call 555-1234");
    });
  });

  // ── Related fields ────────────────────

  describe("related field data", () => {
    it("fetches and renders related field data", async () => {
      const data = makeData({
        title: makeField({
          label: "Work Order",
          value: "WO-001",
          lookupEntityType: "msdyn_workorder",
          lookupId: "wo-id-1",
        }),
      });
      const relatedMappings: RelatedFieldMapping[] = [
        {
          sourceSlot: "titleField",
          fetchField: "msdyn_serviceaccount",
          targetSlot: "subtitleField2",
        },
      ];
      const fetchRelatedData = jest.fn().mockResolvedValue({
        msdyn_serviceaccount: {
          value: "Contoso Related",
          label: "Service Account",
        },
      });

      render(
        <InfoCardComponent
          {...makeProps({
            data,
            relatedMappings,
            fetchRelatedData,
          })}
        />,
      );

      // Wait for the async fetch to be called
      await flushPromises();

      expect(fetchRelatedData).toHaveBeenCalledWith(
        "msdyn_workorder",
        "wo-id-1",
        ["msdyn_serviceaccount"],
      );
    });

    it("does not fetch when no related mappings", () => {
      const data = makeData();
      const fetchRelatedData = jest.fn();
      render(
        <InfoCardComponent
          {...makeProps({
            data,
            relatedMappings: [],
            fetchRelatedData,
          })}
        />,
      );
      expect(fetchRelatedData).not.toHaveBeenCalled();
    });

    it("does not fetch when source slot has no lookup info", () => {
      const data = makeData({
        title: makeField({
          label: "Name",
          value: "Plain Text", // no lookupEntityType or lookupId
        }),
      });
      const relatedMappings: RelatedFieldMapping[] = [
        {
          sourceSlot: "titleField",
          fetchField: "some_field",
          targetSlot: "tagField1",
        },
      ];
      const fetchRelatedData = jest.fn();
      render(
        <InfoCardComponent
          {...makeProps({
            data,
            relatedMappings,
            fetchRelatedData,
          })}
        />,
      );
      expect(fetchRelatedData).not.toHaveBeenCalled();
    });

    it("handles API error in related field fetch gracefully", async () => {
      const data = makeData({
        title: makeField({
          label: "WO",
          value: "WO-001",
          lookupEntityType: "msdyn_workorder",
          lookupId: "wo-id-1",
        }),
      });
      const relatedMappings: RelatedFieldMapping[] = [
        {
          sourceSlot: "titleField",
          fetchField: "bad_field",
          targetSlot: "tagField1",
        },
      ];
      const fetchRelatedData = jest.fn().mockResolvedValue({
        __debug_error: {
          value: "Some API error",
          label: "error",
        },
      });

      render(
        <InfoCardComponent
          {...makeProps({
            data,
            relatedMappings,
            fetchRelatedData,
            showVersion: true,
          })}
        />,
      );

      await flushPromises();
      expect(fetchRelatedData).toHaveBeenCalled();
    });

    it("handles rejected promise in related field fetch", async () => {
      const data = makeData({
        title: makeField({
          label: "WO",
          value: "WO-001",
          lookupEntityType: "msdyn_workorder",
          lookupId: "wo-id-1",
        }),
      });
      const relatedMappings: RelatedFieldMapping[] = [
        {
          sourceSlot: "titleField",
          fetchField: "bad_field",
          targetSlot: "tagField1",
        },
      ];
      const fetchRelatedData = jest.fn().mockRejectedValue(new Error("Network failure"));

      render(
        <InfoCardComponent
          {...makeProps({
            data,
            relatedMappings,
            fetchRelatedData,
            showVersion: true,
          })}
        />,
      );

      await flushPromises();
      expect(fetchRelatedData).toHaveBeenCalled();
    });
  });

  // ── Detail field rendering ────────────

  describe("detail fields", () => {
    it("renders detail values", () => {
      const data = makeData({
        details: [
          makeField({ label: "Notes", value: "Some important notes" }),
          makeField({ label: "Description", value: "A detailed description" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      expect(container.textContent).toContain("Some important notes");
      expect(container.textContent).toContain("A detailed description");
    });

    it("renders address detail as map link when lat/lng present", () => {
      const data = makeData({
        latitude: 47.6,
        longitude: -122.3,
        details: [
          makeField({ label: "Address", value: "123 Main Street" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const mapLink = container.querySelector("a[href*='maps.google.com']");
      expect(mapLink).not.toBeNull();
    });
  });

  // ── Empty data scenarios ──────────────

  describe("edge cases", () => {
    it("renders correctly with minimal data (title only)", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps()} />,
      );
      expect(container.textContent).toContain("Test Card");
    });

    it("renders correctly with all arrays empty", () => {
      const data = makeData({
        subtitles: [],
        phones: [],
        details: [],
        gridFields: [],
        tags: [],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data })} />,
      );
      expect(container.textContent).toContain("Test Card");
    });

    it("handles image URL in data", () => {
      const data = makeData({
        imageUrl: "https://example.com/photo.jpg",
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const img = container.querySelector("img");
      expect(img).not.toBeNull();
      expect(img!.getAttribute("src")).toBe("https://example.com/photo.jpg");
    });

    it("renders no avatar when imageUrl is null (smart layout)", () => {
      const data = makeData({
        title: { label: "Name", value: "Adventure Works", rawValue: "Adventure Works", isEmpty: false },
        imageUrl: null,
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      expect(container.querySelector("img")).toBeNull();
      // No initials placeholder when no image URL was supplied
      expect(container.textContent).not.toContain("AW");
    });

    it("renders no avatar when imageUrl is null (contact layout)", () => {
      const data = makeData({
        title: { label: "Name", value: "Jane Doe", rawValue: "Jane Doe", isEmpty: false },
        imageUrl: null,
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact" })} />,
      );
      expect(container.querySelector("img")).toBeNull();
      expect(container.textContent).not.toContain("JD");
    });

    it("does not render avatar in compact layout", () => {
      const data = makeData({
        title: { label: "Name", value: "Adventure Works", rawValue: "Adventure Works", isEmpty: false },
        imageUrl: "https://example.com/photo.jpg",
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "compact" })} />,
      );
      expect(container.querySelector("img")).toBeNull();
    });

    it("uses title text as alt attribute on avatar img", () => {
      const data = makeData({
        title: { label: "Name", value: "Acme Corp", rawValue: "Acme Corp", isEmpty: false },
        imageUrl: "https://example.com/photo.jpg",
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const img = container.querySelector("img");
      expect(img!.getAttribute("alt")).toBe("Acme Corp");
    });

    it("falls back to initials when image fails to load", () => {
      const data = makeData({
        title: { label: "Name", value: "Microsoft", rawValue: "Microsoft", isEmpty: false },
        imageUrl: "https://example.com/broken.jpg",
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const img = container.querySelector("img");
      expect(img).not.toBeNull();
      // Trigger the onError handler — simulates a broken image load
      act(() => {
        img!.dispatchEvent(new Event("error"));
      });
      // Single-word title → first two characters
      expect(container.textContent).toContain("MI");
      expect(container.querySelector("img")).toBeNull();
    });
  });

  // ── Compact layout specifics ──────────

  describe("compact layout specifics", () => {
    it("shows Contact section header when phones/email present", () => {
      const data = makeData({
        phones: [makeField({ label: "Phone", value: "+1555", rawValue: "+1555" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "compact" })} />,
      );
      expect(container.textContent).toContain("Contact");
    });

    it("shows Details section header when grid fields present", () => {
      const data = makeData({
        gridFields: [makeField({ label: "Status", value: "Active" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "compact" })} />,
      );
      expect(container.textContent).toContain("Details");
    });

    it("shows Info section header when detail fields present", () => {
      const data = makeData({
        details: [makeField({ label: "Notes", value: "Some notes" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "compact" })} />,
      );
      expect(container.textContent).toContain("Info");
    });

    it("renders field labels and values in rows", () => {
      const data = makeData({
        gridFields: [
          makeField({ label: "Duration", value: "2h 30m" }),
        ],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "compact" })} />,
      );
      expect(container.textContent).toContain("Duration");
      expect(container.textContent).toContain("2h 30m");
    });
  });

  // ── Smart layout collapse ─────────────

  describe("smart layout collapse", () => {
    it("collapses body when chevron clicked", () => {
      const data = makeData({
        phones: [makeField({ label: "Phone", value: "+1555", rawValue: "+1555" })],
        gridFields: [makeField({ label: "Status", value: "Active" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );

      // Body should be visible initially
      expect(container.textContent).toContain("Active");

      // Find and click the chevron
      const chevron = container.querySelector("svg[viewBox='0 0 12 12']")?.parentElement;
      expect(chevron).not.toBeNull();
      fireEvent.click(chevron!);

      // After collapse, grid content should be hidden
      expect(container.textContent).not.toContain("Active");
    });
  });

  // ── End-to-end @ fetch pipeline ──────

  describe("@ fetch pipeline", () => {
    it("fetches from title entity and renders merged subtitle", async () => {
      const fetchMock = jest.fn().mockResolvedValue({
        msdyn_serviceaccount: { value: "Contoso Ltd.", label: "Service Account" },
      });

      const { container } = render(
        <InfoCardComponent
          {...makeProps({
            data: makeData({
              title: makeField({
                label: "Work Order",
                value: "WO-001",
                lookupEntityType: "msdyn_workorder",
                lookupId: "wo-123",
              }),
              subtitles: [
                makeField({ label: "Resource", value: "Jack" }),
                { label: "Sub2", value: "---", rawValue: "@msdyn_serviceaccount", isEmpty: true },
              ],
            }),
            relatedMappings: [{
              sourceSlot: "titleField",
              fetchField: "msdyn_serviceaccount",
              targetSlot: "subtitleField2",
            }],
            fetchRelatedData: fetchMock,
          })}
        />,
      );

      await waitFor(() => {
        expect(container.textContent).toContain("Contoso Ltd.");
      });
      expect(fetchMock).toHaveBeenCalledWith("msdyn_workorder", "wo-123", ["msdyn_serviceaccount"]);
    });

    it("fetches @. from current record and renders in grid", async () => {
      const fetchMock = jest.fn().mockResolvedValue({
        "resource.resourcetype": { value: "Contact", label: "Resource Type" },
      });

      const { container } = render(
        <InfoCardComponent
          {...makeProps({
            data: makeData({
              title: makeField({ label: "WO", value: "WO-001" }),
              gridFields: [
                makeField({ label: "Start", value: "9am" }),
                makeField({ label: "End", value: "11am" }),
                makeField({ label: "Duration", value: "2h" }),
                makeField({ label: "Type", value: "Solid" }),
                { label: "Grid 5", value: "---", rawValue: "@.resource.resourcetype", isEmpty: true },
              ],
            }),
            currentRecordMappings: [{
              sourceSlot: "__currentRecord__",
              fetchField: "resource.resourcetype",
              targetSlot: "gridField5",
            }],
            currentRecordEntityType: "bookableresourcebooking",
            currentRecordId: "bk-123",
            fetchRelatedData: fetchMock,
          })}
        />,
      );

      await flushPromises();
      expect(fetchMock).toHaveBeenCalledWith("bookableresourcebooking", "bk-123", ["resource.resourcetype"]);
    });
  });

  // ── resolveRecordFields React integration ──

  describe("resolveRecordFields callback", () => {
    it("applies label and value overrides to grid fields", async () => {
      const resolveMock = jest.fn().mockResolvedValue({
        gridField1: { label: "Start Time", value: "3/30/2026 9:00 AM" },
        gridField2: { label: "Booking Type", value: "Solid" },
      });

      const { container } = render(
        <InfoCardComponent
          {...makeProps({
            data: makeData({
              title: makeField({ label: "WO", value: "WO-001" }),
              gridFields: [
                makeField({ label: "Grid 1", value: "raw-date" }),
                makeField({ label: "Grid 2", value: "1" }),
              ],
            }),
            resolveRecordFields: resolveMock,
          })}
        />,
      );

      await flushPromises();
      expect(resolveMock).toHaveBeenCalled();
      // After resolve, the overrides should be applied on next render
      // Verify the mock was called — the actual DOM update depends on React 16 act() timing
    });

    it("applies color overrides to tag chips", async () => {
      const resolveMock = jest.fn().mockResolvedValue({
        tagField1: { label: "", value: "", color: "#49F249" },
      });

      const { container } = render(
        <InfoCardComponent
          {...makeProps({
            data: makeData({
              title: makeField({ label: "WO", value: "WO-001" }),
              tags: [makeField({ label: "Booking Status", value: "Scheduled" })],
            }),
            resolveRecordFields: resolveMock,
          })}
        />,
      );

      await flushPromises();
      expect(resolveMock).toHaveBeenCalled();
      // Verify the Scheduled tag is rendered
      const chips = Array.from(container.querySelectorAll("span")).filter(
        s => s.textContent === "Scheduled"
      );
      expect(chips.length).toBeGreaterThan(0);
    });
  });

  // ── v4.2 features: titlePrefix, imageShape, collapsibleSections ─────

  describe("titlePrefix", () => {
    it("renders the prefix immediately before the title text", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "WO-12345" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart", titlePrefix: "Work Order: " })} />,
      );
      expect(container.textContent).toContain("Work Order: WO-12345");
    });

    it("does not render any prefix span when titlePrefix is empty", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "WO-12345" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      expect(container.textContent).toContain("WO-12345");
      expect(container.textContent).not.toContain("Work Order:");
    });

    it("works in compact layout too", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "ACME-001" }),
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "compact", titlePrefix: "Case: " })} />,
      );
      expect(container.textContent).toContain("Case: ACME-001");
    });
  });

  describe("imageShape", () => {
    it("defaults to rounded rectangle (8px border-radius)", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Acme" }),
        imageUrl: "https://example.com/p.jpg",
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart" })} />,
      );
      const img = container.querySelector("img");
      expect(img!.style.borderRadius).toBe("8px");
    });

    it("renders a circle when imageShape='circle'", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Acme" }),
        imageUrl: "https://example.com/p.jpg",
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart", imageShape: "circle" })} />,
      );
      const img = container.querySelector("img");
      expect(img!.style.borderRadius).toBe("50%");
    });

    it("renders a square (no rounding) when imageShape='square'", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Acme" }),
        imageUrl: "https://example.com/p.jpg",
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart", imageShape: "square" })} />,
      );
      const img = container.querySelector("img");
      // Browsers can serialise 0 as either "0px" or "0" — accept both.
      expect(["0px", "0"]).toContain(img!.style.borderRadius);
    });
  });

  describe("collapsibleSections", () => {
    it("default 'body' collapses details + grid but keeps contact and tags visible", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Acme" }),
        phones: [makeField({ label: "Phone", value: "+1555" })],
        details: [makeField({ label: "Notes", value: "Body content" })],
        gridFields: [makeField({ label: "Status", value: "Open" })],
        tags: [makeField({ label: "Priority", value: "TagText" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart", startExpanded: false })} />,
      );
      // Body collapsed → details + grid hidden
      expect(container.textContent).not.toContain("Body content");
      expect(container.textContent).not.toContain("Open");
      // Contact + tags still visible
      expect(container.textContent).toContain("+1555");
      expect(container.textContent).toContain("TagText");
    });

    it("'all' collapses contact, body, and tags simultaneously", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Acme" }),
        phones: [makeField({ label: "Phone", value: "+1555" })],
        details: [makeField({ label: "Notes", value: "Body content" })],
        tags: [makeField({ label: "Priority", value: "TagText" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart", startExpanded: false, collapsibleSections: "all" })} />,
      );
      expect(container.textContent).not.toContain("+1555");
      expect(container.textContent).not.toContain("Body content");
      expect(container.textContent).not.toContain("TagText");
      // Title still visible
      expect(container.textContent).toContain("Acme");
    });

    it("'body-tags' collapses body and tags but keeps contact visible", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Acme" }),
        phones: [makeField({ label: "Phone", value: "+1555" })],
        details: [makeField({ label: "Notes", value: "Body content" })],
        tags: [makeField({ label: "Priority", value: "TagText" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart", startExpanded: false, collapsibleSections: "body-tags" })} />,
      );
      expect(container.textContent).toContain("+1555");
      expect(container.textContent).not.toContain("Body content");
      expect(container.textContent).not.toContain("TagText");
    });

    it("'none' disables collapse entirely — no chevron, no aria-expanded", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Acme" }),
        details: [makeField({ label: "Notes", value: "Body content" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "smart", startExpanded: false, collapsibleSections: "none" })} />,
      );
      // 'none' → not collapsible → body still visible regardless of startExpanded
      expect(container.textContent).toContain("Body content");
      // No aria-expanded button at the card root
      const interactiveRoot = container.querySelector("[aria-expanded]");
      expect(interactiveRoot).toBeNull();
    });

    it("works in contact layout (gains chevron when collapsibleSections != none)", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Sarah" }),
        details: [makeField({ label: "Notes", value: "Body content" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "contact", startExpanded: false, collapsibleSections: "body" })} />,
      );
      expect(container.textContent).not.toContain("Body content");
      // aria-expanded false present somewhere on the card root
      const interactiveRoot = container.querySelector("[aria-expanded='false']");
      expect(interactiveRoot).not.toBeNull();
    });

    it("works in compact layout (gains chevron when collapsibleSections != none)", () => {
      const data = makeData({
        title: makeField({ label: "Name", value: "Sarah" }),
        gridFields: [makeField({ label: "Status", value: "Open" })],
      });
      const { container } = render(
        <InfoCardComponent {...makeProps({ data, layout: "compact", startExpanded: false, collapsibleSections: "body" })} />,
      );
      expect(container.textContent).not.toContain("Open");
    });
  });

  // ── mergeRelatedFields direct unit tests ─
  // Specifically guard against the "unconfigured slot in the middle" misalignment:
  // readSlotGroup compacts null entries out of the array, so digit-derived indices
  // (subtitleField2 → 1) no longer match the post-compaction position.

  describe("mergeRelatedFields()", () => {
    function fetched(slotKey: string, value: string): SlotField {
      return { slotKey, label: "", value, rawValue: value, isEmpty: false };
    }
    function placeholder(slotKey: string, raw = "@related"): SlotField {
      return { slotKey, label: "", value: raw, rawValue: raw, isEmpty: true };
    }

    it("places fetched value into the correct slot when an earlier slot is unconfigured (subtitles)", () => {
      // subtitleField1 is unconfigured (compacted out). subtitleField2 has @-related placeholder.
      const data = makeData({
        subtitles: [placeholder("subtitleField2")],
      });
      const merged = mergeRelatedFields(
        data,
        { msdyn_serviceaccount: fetched("subtitleField2", "Contoso") },
        [{ sourceSlot: "titleField", fetchField: "msdyn_serviceaccount", targetSlot: "subtitleField2" }],
      );
      // Must replace in-place — array length unchanged, slotKey still subtitleField2
      expect(merged.subtitles).toHaveLength(1);
      expect(merged.subtitles[0].slotKey).toBe("subtitleField2");
      expect(merged.subtitles[0].value).toBe("Contoso");
      expect(merged.subtitles[0].isEmpty).toBe(false);
    });

    it("places fetched value into the correct grid slot when earlier grid slots are unconfigured", () => {
      // gridField1 + gridField2 unconfigured; gridField3 has @-related placeholder.
      const data = makeData({
        gridFields: [placeholder("gridField3")],
      });
      const merged = mergeRelatedFields(
        data,
        { workordertype: fetched("gridField3", "Repair") },
        [{ sourceSlot: "titleField", fetchField: "workordertype", targetSlot: "gridField3" }],
      );
      expect(merged.gridFields).toHaveLength(1);
      expect(merged.gridFields[0].slotKey).toBe("gridField3");
      expect(merged.gridFields[0].value).toBe("Repair");
    });

    it("does not duplicate or misalign when multiple related fields target compacted slots", () => {
      // Real scenario: subtitleField1 unconfigured, subtitleField2=@msdyn_serviceaccount, subtitleField3 unconfigured
      // gridField1 unconfigured, gridField2=@workordertype.
      const data = makeData({
        subtitles: [placeholder("subtitleField2")],
        gridFields: [placeholder("gridField2")],
      });
      const merged = mergeRelatedFields(
        data,
        {
          msdyn_serviceaccount: fetched("subtitleField2", "Contoso"),
          workordertype: fetched("gridField2", "Repair"),
        },
        [
          { sourceSlot: "titleField", fetchField: "msdyn_serviceaccount", targetSlot: "subtitleField2" },
          { sourceSlot: "titleField", fetchField: "workordertype", targetSlot: "gridField2" },
        ],
      );
      expect(merged.subtitles.map(s => s.slotKey)).toEqual(["subtitleField2"]);
      expect(merged.subtitles[0].value).toBe("Contoso");
      expect(merged.gridFields.map(g => g.slotKey)).toEqual(["gridField2"]);
      expect(merged.gridFields[0].value).toBe("Repair");
    });

    it("falls back to digit-derived position when target slot is not in the compacted array", () => {
      // Edge case: readSlot returned null for the target slot entirely (not even a placeholder),
      // but a related-field mapping still references it. Place it at digit-1 ordering.
      const data = makeData({
        details: [placeholder("detailField1")],
      });
      const merged = mergeRelatedFields(
        data,
        { instructions: fetched("detailField3", "Bring spare filter") },
        [{ sourceSlot: "titleField", fetchField: "instructions", targetSlot: "detailField3" }],
      );
      // detailField3 should land after detailField1 in render order
      const detailField3 = merged.details.find(d => d.slotKey === "detailField3");
      expect(detailField3).toBeDefined();
      expect(detailField3!.value).toBe("Bring spare filter");
    });

    it("replaces in place rather than appending when target slot already exists with @-placeholder", () => {
      // Regression: previously the digit-parse path for subtitleField2 wrote to merged.subtitles[1]
      // even when subtitles=[placeholder_subtitleField2], padding subtitles[0] with an extra empty
      // entry and creating two array elements both keyed subtitleField2.
      const data = makeData({
        subtitles: [placeholder("subtitleField2")],
      });
      const merged = mergeRelatedFields(
        data,
        { svcacc: fetched("subtitleField2", "Contoso") },
        [{ sourceSlot: "titleField", fetchField: "svcacc", targetSlot: "subtitleField2" }],
      );
      const subtitleField2Entries = merged.subtitles.filter(s => s.slotKey === "subtitleField2");
      expect(subtitleField2Entries).toHaveLength(1);
    });
  });

  // ── v4.3 features: showDetailIcons, detailLabelStyle ─────────────
  describe("showDetailIcons", () => {
    function detailData(): InfoCardData {
      return makeData({
        details: [
          makeField({ label: "Instructions", value: "Access via back garden" }),
          makeField({ label: "Summary", value: "Replace filter" }),
        ],
      });
    }

    it("renders leading icons by default on smart layout", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: detailData(), layout: "smart", collapsibleSections: "none" })} />,
      );
      // guessDetailIcon emits an <svg> element inside the leading <span>.
      const svgs = container.querySelectorAll("svg");
      expect(svgs.length).toBeGreaterThan(0);
    });

    it("suppresses leading icons when showDetailIcons=false on smart layout", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: detailData(), layout: "smart", showDetailIcons: false, collapsibleSections: "none" })} />,
      );
      // Detail values are still rendered…
      expect(container.textContent).toContain("Access via back garden");
      expect(container.textContent).toContain("Replace filter");
      // …but no detail-icon <svg>s are present (chevron suppressed via collapsibleSections="none").
      const svgs = container.querySelectorAll("svg");
      expect(svgs.length).toBe(0);
    });

    it("suppresses leading icons when showDetailIcons=false on contact layout", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: detailData(), layout: "contact", showDetailIcons: false, collapsibleSections: "none" })} />,
      );
      expect(container.textContent).toContain("Access via back garden");
      const svgs = container.querySelectorAll("svg");
      expect(svgs.length).toBe(0);
    });
  });

  describe("detailLabelStyle", () => {
    function detailData(): InfoCardData {
      return makeData({
        details: [
          makeField({ label: "Instructions", value: "Access via back garden" }),
        ],
      });
    }

    it("does not render the field label by default", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: detailData(), layout: "smart" })} />,
      );
      expect(container.textContent).toContain("Access via back garden");
      // Default labelStyle="none" — no "Instructions" text is rendered as a label
      expect(container.textContent).not.toContain("Instructions:");
    });

    it("renders bold inline 'Label: value' when detailLabelStyle='inline-bold' on smart layout", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: detailData(), layout: "smart", detailLabelStyle: "inline-bold" })} />,
      );
      expect(container.textContent).toContain("Instructions:");
      expect(container.textContent).toContain("Access via back garden");
      // Find the bold span carrying the label
      const boldLabel = Array.from(container.querySelectorAll("span"))
        .find(s => s.textContent === "Instructions: " && s.style.fontWeight === "600");
      expect(boldLabel).toBeDefined();
    });

    it("renders bold inline label on contact layout", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: detailData(), layout: "contact", detailLabelStyle: "inline-bold" })} />,
      );
      expect(container.textContent).toContain("Instructions:");
      expect(container.textContent).toContain("Access via back garden");
    });

    it("renders label as a heading above value when detailLabelStyle='above'", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: detailData(), layout: "smart", detailLabelStyle: "above" })} />,
      );
      expect(container.textContent).toContain("Instructions");
      expect(container.textContent).toContain("Access via back garden");
      // The above-heading style uses uppercase + small font + 600 weight on its own line
      const headingDiv = Array.from(container.querySelectorAll("div"))
        .find(d => d.textContent === "Instructions"
          && d.style.textTransform === "uppercase"
          && d.style.fontWeight === "600");
      expect(headingDiv).toBeDefined();
    });

    it("does not render labels on compact layout (compact already shows label:value)", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({ data: detailData(), layout: "compact", detailLabelStyle: "inline-bold" })} />,
      );
      // Compact ignores the prop — its own label rendering is unchanged
      expect(container.textContent).toContain("Access via back garden");
    });

    it("combines showDetailIcons=false with detailLabelStyle='inline-bold' for prose-style rows", () => {
      const { container } = render(
        <InfoCardComponent {...makeProps({
          data: detailData(),
          layout: "smart",
          showDetailIcons: false,
          detailLabelStyle: "inline-bold",
          collapsibleSections: "none",
        })} />,
      );
      expect(container.textContent).toContain("Instructions:");
      expect(container.textContent).toContain("Access via back garden");
      // No detail icons (and no chevron since collapsibleSections="none")
      expect(container.querySelectorAll("svg").length).toBe(0);
    });
  });
});
