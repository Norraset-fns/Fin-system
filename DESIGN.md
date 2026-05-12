---
name: FinSystem
description: ระบบบันทึกรายรับรายจ่าย และ ใบเสร็จ
colors:
  primary: "#4f46e5"
  success: "#10b981"
  danger: "#f43f5e"
  warning: "#d97706"
  neutral-bg: "#f8fafc"
  neutral-text: "#1e293b"
  sidebar-bg: "#1e1b4b"
typography:
  display:
    fontFamily: "Prompt, sans-serif"
    fontWeight: 700
    fontSize: "1.875rem"
  body:
    fontFamily: "Prompt, sans-serif"
    fontWeight: 400
    fontSize: "1rem"
  label:
    fontFamily: "Prompt, sans-serif"
    fontWeight: 600
    fontSize: "0.75rem"
    letterSpacing: "0.1em"
rounded:
  sm: "0.5rem"
  md: "0.75rem"
  lg: "1rem"
  xl: "1.5rem"
spacing:
  xs: "0.5rem"
  sm: "1rem"
  md: "1.5rem"
  lg: "2rem"
components:
  button-primary:
    backgroundColor: "{colors.primary}"
    textColor: "#ffffff"
    rounded: "{rounded.md}"
    padding: "0.75rem 1rem"
  card:
    backgroundColor: "#ffffff"
    rounded: "{rounded.xl}"
    padding: "2rem"
---

# Design System: FinSystem

## 1. Overview

**Creative North Star: "The Precision Ledger"**

A financial management interface designed for clarity, speed, and accuracy. The system uses a high-contrast Indigo theme to convey professional trust, balanced with soft Slate neutrals to keep the focus on financial data. It rejects the clutter of traditional accounting software in favor of a modern, "app-shell" architecture that feels efficient on both desktop and mobile.

**Key Characteristics:**
- **Indigo Authority**: Deep indigo surfaces provide structure and hierarchy.
- **Color-Coded Feedback**: Semantic use of Emerald (Income) and Rose (Expense) for instant data recognition.
- **Rounded Precision**: Consistent 12px-16px corner radii soften the professional tone, making it feel modern and accessible.

## 2. Colors

The palette is technically grounded in Tailwind's Slate and Indigo scales, used semantically to differentiate status.

### Primary
- **Indigo 600** (#4f46e5): The brand anchor. Used for primary actions, navigation, and active states.

### Success
- **Emerald 500** (#10b981): Represents income, positive growth, and successful operations.

### Danger
- **Rose 500** (#f43f5e): Represents expenses, critical errors, and destructive actions (voiding).

### Neutral
- **Slate 50** (#f8fafc): Main background to reduce eye strain.
- **Slate 800** (#1e293b): Primary text for high readability.
- **Indigo 950** (#1e1b4b): Sidebar background, providing strong lateral containment.

**The Semantic Meaning Rule.** Color is never used decoratively. If a text or icon is Emerald, it MUST relate to "Income" or "Success". If it is Rose, it MUST relate to "Expense" or "Warning".

## 3. Typography

**Display & Body Font:** Prompt (Google Fonts)

**Character:** A modern, geometric Thai-friendly sans-serif that balances technical precision with a friendly, readable curve.

### Hierarchy
- **Display** (700, 1.875rem): Page titles and key financial summaries.
- **Headline** (600, 1.25rem): Section headers and modal titles.
- **Body** (400, 1rem): General content, table data, and form labels.
- **Label** (600, 0.75rem, uppercase, 0.1em tracking): Helper text, table headers, and metadata.

## 4. Elevation

The system uses a layered elevation strategy with subtle shadows to define depth without cluttering the interface.

### Shadow Vocabulary
- **Shadow Sm**: Used for cards and secondary buttons to provide a slight lift from the Slate background.
- **Shadow Xl**: Used for modals and floating headers to create a clear visual break from the main content.

**The Flat-Surface Rule.** All data containers (tables, lists) remain flat at rest. Elevation is only applied to interactive components (buttons on hover) or prioritized surfaces (modals).

## 5. Components

### Buttons
- **Shape:** Rounded-xl (0.75rem).
- **Primary:** Indigo-600 with white text. Bold weight.
- **State:** Hover shifts background slightly darker; focus adds a 2px indigo ring.

### Cards
- **Style:** White background, 1px Slate-100 border, rounded-2xl (1rem).
- **Padding:** 2rem (32px) for ample whitespace.

### Inputs
- **Style:** Slate-50 background, Slate-200 border, rounded-xl.
- **Focus:** Border shifts to Indigo-500 with a subtle ring.

### Navigation (Sidebar)
- **Style:** Indigo-900/950 background with Indigo-400 icons. Active items use Indigo-800 background with white text.

## 6. Do's and Don'ts

### Do:
- **Do** use `rounded-xl` for all primary containers and buttons.
- **Do** maintain a minimum of 2rem padding in main content cards.
- **Do** use Emerald for income and Rose for expenses consistently.

### Don't:
- **Don't** use border-left greater than 1px as a colored stripe on cards.
- **Don't** use cluttered enterprise-style tables; use Slate-50 headers and ample row padding.
- **Don't** use pure black (#000) or pure white (#fff) for large surfaces; use Slate-800 and Slate-50.
