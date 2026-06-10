---
name: Precision Logistics
colors:
  surface: '#f8f9ff'
  surface-dim: '#cbdbf5'
  surface-bright: '#f8f9ff'
  surface-container-lowest: '#ffffff'
  surface-container-low: '#eff4ff'
  surface-container: '#e5eeff'
  surface-container-high: '#dce9ff'
  surface-container-highest: '#d3e4fe'
  on-surface: '#0b1c30'
  on-surface-variant: '#45464d'
  inverse-surface: '#213145'
  inverse-on-surface: '#eaf1ff'
  outline: '#76777d'
  outline-variant: '#c6c6cd'
  surface-tint: '#565e74'
  primary: '#000000'
  on-primary: '#ffffff'
  primary-container: '#131b2e'
  on-primary-container: '#7c839b'
  inverse-primary: '#bec6e0'
  secondary: '#0051d5'
  on-secondary: '#ffffff'
  secondary-container: '#316bf3'
  on-secondary-container: '#fefcff'
  tertiary: '#000000'
  on-tertiary: '#ffffff'
  tertiary-container: '#271901'
  on-tertiary-container: '#98805d'
  error: '#ba1a1a'
  on-error: '#ffffff'
  error-container: '#ffdad6'
  on-error-container: '#93000a'
  primary-fixed: '#dae2fd'
  primary-fixed-dim: '#bec6e0'
  on-primary-fixed: '#131b2e'
  on-primary-fixed-variant: '#3f465c'
  secondary-fixed: '#dbe1ff'
  secondary-fixed-dim: '#b4c5ff'
  on-secondary-fixed: '#00174b'
  on-secondary-fixed-variant: '#003ea8'
  tertiary-fixed: '#fcdeb5'
  tertiary-fixed-dim: '#dec29a'
  on-tertiary-fixed: '#271901'
  on-tertiary-fixed-variant: '#574425'
  background: '#f8f9ff'
  on-background: '#0b1c30'
  surface-variant: '#d3e4fe'
typography:
  display-lg:
    fontFamily: Inter
    fontSize: 32px
    fontWeight: '700'
    lineHeight: 40px
    letterSpacing: -0.02em
  headline-md:
    fontFamily: Inter
    fontSize: 24px
    fontWeight: '600'
    lineHeight: 32px
    letterSpacing: -0.01em
  title-sm:
    fontFamily: Inter
    fontSize: 18px
    fontWeight: '600'
    lineHeight: 24px
  body-md:
    fontFamily: Inter
    fontSize: 14px
    fontWeight: '400'
    lineHeight: 20px
  body-sm:
    fontFamily: Inter
    fontSize: 13px
    fontWeight: '400'
    lineHeight: 18px
  label-caps:
    fontFamily: Inter
    fontSize: 11px
    fontWeight: '700'
    lineHeight: 16px
    letterSpacing: 0.05em
  data-mono:
    fontFamily: JetBrains Mono
    fontSize: 13px
    fontWeight: '500'
    lineHeight: 16px
rounded:
  sm: 0.25rem
  DEFAULT: 0.5rem
  md: 0.75rem
  lg: 1rem
  xl: 1.5rem
  full: 9999px
spacing:
  base: 4px
  xs: 8px
  sm: 12px
  md: 16px
  lg: 24px
  xl: 32px
  gutter: 20px
  sidebar_width: 260px
---

## Brand & Style

The design system is engineered for high-density operational environments where clarity and speed of cognition are paramount. The brand personality is **Reliable, Efficient, and Expert**, focusing on a "no-friction" user experience for logistics coordinators.

The aesthetic follows a **Corporate / Modern** style with a focus on functional minimalism. It prioritizes data legibility and systematic organization over decorative elements. The UI evokes an emotional response of being "in control" and "well-informed" through a structured, white-label professional appearance.

- **Primary Motif:** Information density balanced with ample whitespace.
- **Visual Tone:** Technical, trustworthy, and precise.
- **Target Audience:** Warehouse managers, fleet dispatchers, and logistics analysts.

## Colors

The palette is rooted in a professional "Navy and Slate" foundation to communicate stability and authority. 

- **Primary (#0F172A):** Used for navigation, high-level headers, and core brand moments.
- **Secondary (#2563EB):** The functional action color, used for buttons, active states, and primary data callouts.
- **Neutrals:** A range of Cool Grays are used to create structural hierarchy. `#F8FAFC` is the primary canvas, while `#F1F5F9` differentiates secondary containers and sidebars.
- **Semantic Colors:** Critical for logistics. **Success (Green)** for completed deliveries; **Warning (Amber)** for high workloads or delayed processing; **Danger (Red)** for capacity breaches or urgent fleet alerts.

## Typography

The typography system utilizes **Inter** for its exceptional legibility in data-heavy interfaces. It is complemented by **JetBrains Mono** for specific technical data points like tracking numbers, coordinates, and timestamps to ensure character distinction.

- **Headlines:** Use tight letter spacing and heavier weights to anchor sections.
- **Body:** Standardized at 14px for optimal density on desktop displays.
- **Labels:** Uppercase labels with slight tracking are used for table headers and category descriptors.
- **Data Mono:** Reserved for unique identifiers (SKUs, Order IDs) to prevent reading errors between '0' and 'O' or '1' and 'l'.

## Layout & Spacing

This design system uses a **Fixed-Fluid Hybrid Grid**. The navigation sidebar is fixed at 260px, while the main content area utilizes a 12-column fluid grid that adapts to the viewport width.

- **Rhythm:** An 8px linear scale (with a 4px half-step for tight components) governs all padding and margins.
- **Margins:** Main page margins are set to 32px (xl) to provide visual "breathing room" against the dense data tables.
- **Gaps:** Standard component spacing is 20px to ensure clear separation between dashboard widgets.
- **Mobile Reflow:** On mobile devices, the 12-column desktop grid collapses into a single-column stack. Charts and maps maintain a minimum height of 300px to ensure usability.

## Elevation & Depth

The design system utilizes **Tonal Layering** combined with **Ambient Shadows** to define hierarchy without clutter.

- **Level 0 (Background):** `#F8FAFC` - The lowest layer.
- **Level 1 (Cards/Surface):** White (#FFFFFF) surfaces with a subtle `0px 1px 3px rgba(15, 23, 42, 0.08)` shadow. This provides a "paper-thin" lift.
- **Level 2 (Active/Hover):** Enhanced shadow `0px 10px 15px -3px rgba(15, 23, 42, 0.1)` used for active cards or items being dragged in the logistics queue.
- **Overlays:** Modals and dropdowns use a high-diffusion shadow with a 10% opacity primary color tint to maintain brand consistency even in the shadows.

## Shapes

The shape language is **Rounded**, striking a balance between modern friendliness and professional structure.

- **Standard Elements:** Buttons, input fields, and small tags use a **0.5rem (8px)** radius.
- **Containers:** Dashboard cards and map overlays use a **1rem (16px)** radius to create a soft, contained look.
- **Indicators:** Status "pills" (e.g., "In Transit") use a fully rounded/stadium shape to distinguish them from interactive buttons.

## Components

### Buttons
- **Primary:** Solid `#2563EB` with white text. 8px corner radius.
- **Secondary:** Ghost style with `#64748B` border and text.
- **States:** Hover states should darken the background by 10%.

### Cards & KPIs
- KPI indicators should feature a large `display-lg` value, a `label-caps` title, and a small trend sparkline.
- Cards must include a standard 16px internal padding.

### Lists & Tables
- **Expandable Lists:** Use a chevron-right icon that rotates 90 degrees on expansion. Row height: 48px.
- **Zebra Striping:** Use `#F1F5F9` on even rows for high-density tables to improve horizontal scanning.

### Input Fields
- Border-based inputs using a 1px `#E2E8F0` stroke. On focus, the stroke becomes `#2563EB` with a 2px outer glow.

### Central Map Interface
- The map should use a "Light Gray" vector style (e.g., Positron/Alidade) to ensure primary and status colors (Blue/Amber/Red) pop against the geography. 
- Map markers are circular with 2px white borders.