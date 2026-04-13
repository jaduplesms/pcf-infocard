/**
 * Component tests for InfoCardComponent (InfoCard.tsx).
 *
 * Covers rendering, layout modes, field visibility, action buttons,
 * theme colors, version badge, lookup navigation, and related fields.
 */

import * as React from "react";
import { render, screen, fireEvent, waitFor } from "@testing-library/react";
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
    theme: overrides.theme ?? defaultTheme,
    version: overrides.version ?? "2.4.7",
    relatedMappings: overrides.relatedMappings ?? [],
    fetchRelatedData: overrides.fetchRelatedData,
    onOpenRecord: overrides.onOpenRecord,
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
});
