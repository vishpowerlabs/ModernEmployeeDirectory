# 👥 Modern Employee Directory — SharePoint SPFx Web Part

[![SPFx](https://img.shields.io/badge/SPFx-1.18+-purple?style=flat-square)](https://aka.ms/spfx)
[![Node.js](https://img.shields.io/badge/Node.js-18.x-green?style=flat-square)](https://nodejs.org)
[![React](https://img.shields.io/badge/React-17-blue?style=flat-square)](https://reactjs.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)](./LICENSE)
[![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-v1.0%20%2B%20Beta-00a1f1?style=flat-square)](https://learn.microsoft.com/en-us/graph/overview)

> A feature-rich, open-source SPFx Employee Directory for SharePoint Online and Microsoft Teams — powered by Microsoft Graph API. No SaaS. No external database. No extra licences.

📖 **Full documentation & setup guide:** [wrvishnu.com/modern-employee-directory-sharepoint](https://www.wrvishnu.com/modern-employee-directory-sharepoint/)

---

## 📸 Screenshots

| Grid View | List View |
|-----------|-----------|
| ![Grid View](https://raw.githubusercontent.com/vishpowerlabs/ModernEmployeeDirectory/main/emp_dir_grid_view.png) | ![List View](https://raw.githubusercontent.com/vishpowerlabs/ModernEmployeeDirectory/main/emp_dir_list_view.png) |

| Grid View — Selected Card | List View — Account Menu |
|--------------------------|--------------------------|
| ![Selected Card](https://raw.githubusercontent.com/vishpowerlabs/ModernEmployeeDirectory/main/emp_dir_grid_selected_card.png) | ![Account Menu](https://raw.githubusercontent.com/vishpowerlabs/ModernEmployeeDirectory/main/emp_dir_list_account_menu.png) |

| Employee Profile & Self-Service Update | Web Part Settings |
|----------------------------------------|-------------------|
| ![Profile & Update](https://raw.githubusercontent.com/vishpowerlabs/ModernEmployeeDirectory/main/emp_dir_profile_and_update.png) | ![Web Part Settings](https://raw.githubusercontent.com/vishpowerlabs/ModernEmployeeDirectory/main/Employee0Directory-WebparSettings.png) |

---

## ✨ Features

### 🔍 Directory & Search
- Real-time employee search by name, title, or department
- A–Z alphabetical name filter
- Multi-field dropdown filters: **Department, Job Title, City, Country**
- Switch between **Grid Card** and **List** layouts

### 🌳 Org Chart
- Visualise management hierarchies using three layouts:
  - **Vertical Tree**
  - **Horizontal Tree**
  - **Compact List**
- Powered by Graph `/manager` and `/directReports` endpoints

### 🏆 Kudos & Hall of Fame
- Employees can send peer recognition with badges and messages
- Top kudos earners surface automatically in a **Hall of Fame**
- Manually pin featured employees via property pane
- Configurable minimum kudos threshold for Hall of Fame eligibility

### ✏️ Self-Service Profile Editing
- Employees update their own profiles directly from the directory
- Editable fields: **Bio, Skills, Interests, Mobile Phone, Job Title, Office Location, Past Projects**
- Writes back to Microsoft 365 via **Graph PATCH `/me`** (beta endpoint)
- Admins control which fields are user-editable via the property pane

### 🚀 Deployment Flexibility
- Runs as a **SharePoint Online web part**
- Deployable as a **Microsoft Teams Tab** or **Teams Personal App**
- Fully themeable via **Fluent UI tokens** — adapts to any SharePoint or Teams theme automatically

---

## 🔧 Microsoft Graph API — Permissions Required

A **Tenant Administrator** must approve the following API permissions in the SharePoint Admin Center:

| Permission | Scope | Purpose |
|---|---|---|
| `User.Read.All` | Delegated | Read all user profiles |
| `User.ReadWrite` | Delegated | Self-service profile updates |
| `People.Read.All` | Delegated | Colleague intelligence |

> **Note:** Self-service profile fields (Skills, Interests) use the **Microsoft Graph Beta endpoint** (`/users/{me}/profile/skills`, `/users/{me}/profile/interests`).

---

## ⚙️ Property Pane Configuration

The web part is organised across **3 property pane pages**:

### Page 1 — Display Settings
| Setting | Description |
|---|---|
| Description Field | Web part label shown in the SharePoint page |
| Container Margin | Outer padding in px |
| Badge Circle Size | Employee avatar circle diameter |
| Badge Font Size | Font size for name badge |
| Profile Layout | `Scrolling`, `Tabbed`, or `Modal Overlay` |
| Org Chart Layout | `Vertical Tree`, `Horizontal Tree`, `Compact List` |
| Homepage Title Font Size | Controls hero heading size |
| Detail Page Title Font Size | Controls profile page heading |
| Section Heading Font Size | Controls user name display size |
| Enable Pagination | Toggle paginated loading |

### Page 2 — Organisation Filters & Kudos
| Setting | Description |
|---|---|
| Filter Type | Filter directory `By Department` or other fields |
| Department Name | Default department scoping |
| Home Page Dropdown Filters | Which filters show on the homepage |
| Enable Kudos | Toggle the kudos & recognition system |
| Min Kudos for Hall of Fame | Threshold to appear in Hall of Fame |
| Select Kudos List | Point to a SharePoint List to store kudos |
| Manually Featured People | Pin specific employees to Hall of Fame |

### Page 3 — Self-Service Settings
| Setting | Description |
|---|---|
| Updatable Profile Fields | Select which fields employees can self-edit |

Available self-editable fields: `Job Title`, `Bio (About Me)`, `Mobile Phone`, `Office Location`, `Skills`, `Interests`, `Past Projects`

---

## 🛠️ Tech Stack

| Technology | Purpose |
|---|---|
| [SPFx 1.18+](https://aka.ms/spfx) | SharePoint Framework scaffolding |
| [React 17](https://reactjs.org) | UI component library |
| [MSGraphClientV3](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/use-msgraph) | Microsoft Graph API client |
| [Fluent UI v8](https://developer.microsoft.com/en-us/fluentui) | Microsoft design system components |
| [react-organizational-chart](https://github.com/bumbu/react-organizational-chart) | Org chart visualisation |
| SharePoint Lists | Kudos storage (no external DB) |

---

## 🚀 Getting Started

### Prerequisites
- Node.js **18.x**
- SharePoint Online tenant (Microsoft 365)
- Tenant Admin access (for API permissions approval)

### 1. Clone the repository

```bash
git clone https://github.com/vishpowerlabs/ModernEmployeeDirectory.git
cd ModernEmployeeDirectory
```

### 2. Install dependencies

```bash
npm install
```

### 3. Trust the developer certificate

```bash
gulp trust-dev-cert
```

### 4. Run locally

```bash
gulp serve
```

### 5. Build for production

```bash
gulp bundle --ship
gulp package-solution --ship
```

This generates the `.sppkg` file in the `sharepoint/solution` folder.

### 6. Deploy to SharePoint

1. Upload the `.sppkg` to your **SharePoint App Catalog**
2. Click **Deploy** and approve tenant-wide deployment
3. A Tenant Admin must approve the **Graph API permissions** in the SharePoint Admin Center → API Access page
4. Add the web part to any SharePoint page

---

## 📋 SharePoint List Setup (Kudos)

The Kudos system requires a SharePoint List with the following columns:

| Column Name | Type | Notes |
|---|---|---|
| `Title` | Single line of text | Auto-generated |
| `FromUser` | Person or Group | Kudos sender |
| `ToUser` | Person or Group | Kudos recipient |
| `BadgeType` | Choice | e.g. Star, Innovator, Team Player |
| `Message` | Multiple lines of text | Recognition message |
| `KudosDate` | Date and Time | Date of recognition |

> After creating the list, point the web part to it via **Property Pane → Page 2 → Select Kudos List**.

---

## 🗺️ Roadmap

Features being considered for future releases — **your feedback shapes this list:**

- [ ] Skill-based employee search & discovery
- [ ] Org chart export to PDF / image
- [ ] Viva Connections card integration
- [ ] Bulk kudos import from CSV
- [ ] Department-level directory scoping
- [ ] Dark mode support

💬 **Have a feature request or found a bug?** Open an [Issue](https://github.com/vishpowerlabs/ModernEmployeeDirectory/issues) or drop a comment on the [blog post](https://www.wrvishnu.com/modern-employee-directory-sharepoint/).

---

## 🤝 Contributing

Contributions are welcome! Whether it's a bug fix, a new feature, or a documentation improvement — all PRs are appreciated.

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature-name`
3. Commit your changes: `git commit -m 'Add: your feature description'`
4. Push to the branch: `git push origin feature/your-feature-name`
5. Open a Pull Request

---

## 📄 Licence

This project is licensed under the **MIT Licence** — see the [LICENSE](./LICENSE) file for details.

Free to use, modify, and distribute. Attribution appreciated but not required.

---

## 👨‍💻 Author

**Vish** — Microsoft Solutions Architect | SharePoint & Power Platform

- 🌐 Blog: [wrvishnu.com](https://www.wrvishnu.com)
- 💼 LinkedIn: [linkedin.com/in/vishnuwr](https://linkedin.com/in/vishnuwr)
- 🐙 GitHub: [github.com/vishpowerlabs/ModernEmployeeDirectory](https://github.com/vishpowerlabs/ModernEmployeeDirectory)

---

## ⭐ Support

If this project helped you, consider giving it a **star** ⭐ — it helps others in the community find it.

Found it useful in your org? Share your experience in the [blog comments](https://www.wrvishnu.com/modern-employee-directory-sharepoint/) — feedback drives the roadmap!
