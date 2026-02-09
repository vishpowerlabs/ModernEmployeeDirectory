# Development Workbench - Full Width Mode

This file contains instructions for enabling full-width mode in the SPFx workbench during development.

## Quick Setup

### Option 1: Browser Console (Easiest)
1. Open the workbench: `gulp serve` or `npm start`
2. Open browser DevTools (F12)
3. Go to Console tab
4. Paste this code and press Enter:

```javascript
const style = document.createElement('style');
style.textContent = `
.p_ZwSiC_hHQBj,
.CanvasComponent,
.LCS,
.CanvasZone,
div[data-automation-id="CanvasZone"],
.CanvasSection {
    max-width: 100% !important;
    width: 100% !important;
}
.CanvasComponent.LCS .CanvasZone {
    max-width: 100% !important;
    width: 100% !important;
    padding: 0 !important;
    margin: 0 auto !important;
}
`;
document.head.appendChild(style);
console.log('✅ Full-width mode enabled!');
```

### Option 2: Browser Extension (Persistent)
1. Install a CSS injection extension like "Stylus" or "User CSS"
2. Create a new style for `https://localhost:4321/*`
3. Copy the contents from `dev-workbench-fullwidth.css`
4. Save and enable

### Option 3: DevTools Overrides
1. Open DevTools → Sources → Overrides
2. Enable Local Overrides
3. Add the CSS from `dev-workbench-fullwidth.css` to the page

## Important Notes

⚠️ **DEVELOPMENT ONLY** - These styles are NOT included in production builds.

⚠️ **Temporary** - Styles will reset when you refresh the page (unless using Option 2).

⚠️ **Workbench Only** - These overrides only affect the local workbench, not SharePoint.

## Files
- `dev-workbench-fullwidth.css` - The CSS file with all overrides
- This README - Instructions for usage
