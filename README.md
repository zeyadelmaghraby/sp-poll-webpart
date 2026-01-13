# SP Poll Web Part

A SharePoint Framework (SPFx) web part that enables interactive polling functionality in SharePoint Online. Users can vote on questions and view real-time results displayed as pie charts.

![SPFx Version](https://img.shields.io/badge/SPFx-1.22.1-green.svg)
![Node.js](https://img.shields.io/badge/Node.js-%3E%3D22.14.0-green.svg)
![TypeScript](https://img.shields.io/badge/TypeScript-5.8.3-blue.svg)
![SharePoint Online](https://img.shields.io/badge/SharePoint%20Online-Compatible-green.svg)

---

## Table of Contents

- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Development](#development)
- [Building for Production](#building-for-production)
- [Deployment](#deployment)
- [SharePoint Lists Setup](#sharepoint-lists-setup)
- [Configuration](#configuration)
- [Usage](#usage)
- [Troubleshooting](#troubleshooting)
- [Project Structure](#project-structure)

---

## Features

- ✅ Interactive poll questions with up to 4 options
- ✅ Real-time pie chart visualization of results
- ✅ One vote per user enforcement
- ✅ Responsive design with Fluent UI components
- ✅ SharePoint theme integration
- ✅ Works in SharePoint Online, Teams, and Viva Connections
- ✅ Automatic list provisioning on deployment

---

## Prerequisites

Before you begin, ensure you have the following installed:

| Requirement | Version | Installation |
|-------------|---------|--------------|
| **Node.js** | ≥22.14.0 and <23.0.0 | [Download](https://nodejs.org/) |
| **npm** | ≥10.0.0 | Included with Node.js |
| **SharePoint Online** | - | Microsoft 365 tenant |

### Verify Node.js Version

```bash
node --version
# Should output v22.x.x
```

---

## Installation

1. **Clone or download the repository**

2. **Install dependencies**
   ```bash
   cd sp-poll-webpart
   npm install
   ```

3. **Trust the development certificate** (first time only)
   ```bash
   # Open https://localhost:4321 in browser and accept the certificate
   ```

---

## Development

### Start Development Server

```bash
npm run start
```

This will:
- Build the project in watch mode
- Start webpack dev server at `https://localhost:4321`
- Enable hot module replacement (HMR)

### Test in SharePoint Workbench

1. Start the dev server (`npm run start`)
2. Navigate to your SharePoint site workbench:
   ```
   https://[your-tenant].sharepoint.com/sites/[your-site]/_layouts/15/workbench.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true
   ```
3. Click **+** and add the **React Poll** web part

### Available Scripts

| Command | Description |
|---------|-------------|
| `npm run start` | Start development server with watch mode |
| `npm run build` | Build the project |
| `npm run bundle` | Bundle assets for production |
| `npm run bundle:ship` | Bundle assets for production (shipping) |
| `npm run package` | Create .sppkg package |
| `npm run package:ship` | Create production .sppkg package |
| `npm run clean` | Clean build artifacts |
| `npm run lint` | Run ESLint |

---

## Building for Production

### Step 1: Bundle the Solution

```bash
npm run bundle:ship
```

### Step 2: Create the Package

```bash
npm run package:ship
```

This creates the deployment package at:
```
sharepoint/solution/sp-poll-webpart.sppkg
```

### One-Line Build Command

```bash
npm run bundle:ship && npm run package:ship
```

---

## Deployment

### Option 1: Deploy to Tenant App Catalog (Recommended)

1. **Open SharePoint Admin Center**
   - Go to: `https://[your-tenant]-admin.sharepoint.com`
   - Navigate to **More features** → **Apps** → **App Catalog**

2. **Upload the Package**
   - Click **Apps for SharePoint**
   - Upload `sharepoint/solution/sp-poll-webpart.sppkg`
   - Check **"Make this solution available to all sites in the organization"** for tenant-wide deployment
   - Click **Deploy**

3. **Enable on Site (if not tenant-wide)**
   - Go to site: **Site Settings** → **Site Collection Features**
   - Activate **"SP Poll Web Part Feature"****

### Option 2: Deploy to Site Collection App Catalog

1. **Create Site Collection App Catalog** (if not exists)
   - Using SharePoint Admin Center or PowerShell

2. **Upload the Package**
   - Go to site: **Site Contents** → **Apps for SharePoint**
   - Upload `sp-poll-webpart.sppkg`
   - Click **Deploy**

### Verify Deployment

After deployment, the solution will automatically provision the required lists (PollQuestions and PollAnswers) when the feature is activated.

---

## SharePoint Lists Setup

The web part requires two SharePoint lists. These are **automatically created** when you deploy the solution package. However, if you need to create them manually:

### List 1: PollQuestions

**List URL:** `/Lists/PollQuestions`

| Column Name | Type | Required | Description |
|-------------|------|----------|-------------|
| Title | Single line of text | Yes | Default column (can be hidden) |
| Question | Multiple lines of text (Plain) | Yes | The poll question text |
| Option1 | Single line of text | Yes | First answer option |
| Option2 | Single line of text | Yes | Second answer option |
| Option3 | Single line of text | No | Third answer option (optional) |
| Option4 | Single line of text | No | Fourth answer option (optional) |
| IsQuestionActive | Yes/No | Yes | Whether the question is active (Default: Yes) |

#### Create Manually via SharePoint UI:

1. Go to your SharePoint site
2. Click **Settings** (gear icon) → **Site Contents**
3. Click **New** → **List** → **Blank list**
4. Name: `PollQuestions`
5. Click **Create**
6. Add each column using **+ Add column**:
   - **Question**: Multiple lines of text → Plain text → Required
   - **Option1**: Single line of text → Required
   - **Option2**: Single line of text → Required
   - **Option3**: Single line of text → Not required
   - **Option4**: Single line of text → Not required
   - **IsQuestionActive**: Yes/No → Default value: Yes → Required

---

### List 2: PollAnswers

**List URL:** `/Lists/PollAnswers`

| Column Name | Type | Required | Description |
|-------------|------|----------|-------------|
| Title | Single line of text | Yes | Stores the Question ID (reference) |
| UserEmail | Single line of text | Yes | Email of the user who voted |
| Option | Choice | Yes | The selected option |

**Choice Options for Option column:**
- Option1
- Option2
- Option3
- Option4

#### Create Manually via SharePoint UI:

1. Go to your SharePoint site
2. Click **Settings** (gear icon) → **Site Contents**
3. Click **New** → **List** → **Blank list**
4. Name: `PollAnswers`
5. Click **Create**
6. Add each column using **+ Add column**:
   - **UserEmail**: Single line of text → Required
   - **Option**: Choice → Add choices: Option1, Option2, Option3, Option4 → Required

---

### Add Sample Poll Question

After creating the lists, add a test question to **PollQuestions**:

| Column | Sample Value |
|--------|--------------|
| Title | Poll 1 |
| Question | What is your preferred work arrangement? |
| Option1 | Work from Office |
| Option2 | Work from Home |
| Option3 | Hybrid |
| Option4 | Flexible |
| IsQuestionActive | **Yes** |

> ⚠️ **Important:** At least one question must have `IsQuestionActive = Yes` for the web part to display content.

---

## Configuration

### Web Part Properties

The web part automatically detects:
- Current user's email address
- Site URL for list operations

**No manual configuration is required** after deployment.

### serve.json Configuration (Development)

For development, update `config/serve.json` to set your tenant:

```json
{
  "$schema": "https://developer.microsoft.com/json-schemas/spfx-build/spfx-serve.schema.json",
  "port": 4321,
  "https": true,
  "initialPage": "https://[your-tenant].sharepoint.com/sites/[your-site]/_layouts/15/workbench.aspx"
}
```

Or set an environment variable before running:

**Windows (PowerShell):**
```powershell
$env:SPFX_SERVE_TENANT_DOMAIN = "[your-tenant].sharepoint.com"
npm run start
```

**Windows (CMD):**
```cmd
set SPFX_SERVE_TENANT_DOMAIN=[your-tenant].sharepoint.com
npm run start
```

---

## Usage

### Adding the Web Part to a Page

1. Navigate to a SharePoint page
2. Click **Edit** in the top right
3. Click **+** to add a web part
4. Search for **"React Poll"**
5. Click to add the web part to the page
6. Click **Publish** or **Republish**

### How It Works

1. **Active Question**: The web part displays the first active question from the PollQuestions list (ordered by ID)
2. **Voting**: Users select an option and click **Submit**
3. **One Vote Rule**: Each user can vote only once per question (enforced by email)
4. **Results**: After voting, users see a pie chart showing all responses

### Managing Polls

| Action | How To |
|--------|--------|
| **Create a new poll** | Add a new item to PollQuestions list |
| **Activate a poll** | Set `IsQuestionActive = Yes` |
| **Deactivate a poll** | Set `IsQuestionActive = No` |
| **View responses** | Check the PollAnswers list |

> 💡 **Tip:** Keep only one question active at a time for the best user experience.

---

## Troubleshooting

### Common Issues

#### 1. Web Part Not Loading in Workbench

**Problem:** CORS error or certificate issues

**Solution:**
1. Open `https://localhost:4321/temp/build/manifests.js` directly in your browser
2. Accept the security certificate warning
3. Refresh the workbench page

#### 2. "List not found" Error in Console

**Problem:** PollQuestions or PollAnswers lists don't exist

**Solution:**
- Deploy the solution package (auto-creates lists), OR
- Create lists manually (see [SharePoint Lists Setup](#sharepoint-lists-setup))

#### 3. Blank Web Part (No Content)

**Problem:** No active questions in the list

**Solution:**
- Add a question to the PollQuestions list
- Ensure `IsQuestionActive = Yes`

#### 4. Build Errors

**Problem:** Node.js version mismatch

**Solution:**
```bash
node --version  # Must be v22.x.x

# If using nvm (Node Version Manager):
nvm install 22
nvm use 22
```

#### 5. Package Command Fails

**Problem:** Missing bundle step

**Solution:** Always bundle before packaging:
```bash
npm run bundle:ship
npm run package:ship
```

### Enable Debug Mode

Add `?debug=true` to the workbench URL for detailed console logging:
```
https://[tenant].sharepoint.com/.../workbench.aspx?debugManifestsFile=...&debug=true
```

---

## Project Structure

```
sp-poll-webpart/
├── config/                          # Build configuration
│   ├── config.json                  # General build config
│   ├── package-solution.json        # Solution packaging settings
│   ├── rig.json                     # Heft rig configuration
│   ├── sass.json                    # SASS compilation settings
│   ├── serve.json                   # Dev server settings
│   └── typescript.json              # TypeScript settings
├── sharepoint/
│   ├── assets/                      # SharePoint provisioning assets
│   │   ├── elements.xml             # Feature elements (fields, content types, lists)
│   │   ├── schema.xml               # PollAnswers list schema
│   │   └── PollQuestionsSchema.xml  # PollQuestions list schema
│   └── solution/                    # Generated .sppkg output
├── src/
│   ├── apiHooks/
│   │   └── useGetData.ts            # React hook for SharePoint data operations
│   ├── Constants/
│   │   └── Constants.ts             # Field names and list URLs
│   ├── models/
│   │   └── models.ts                # TypeScript interfaces
│   └── webparts/
│       └── reactPoll/
│           ├── ReactPollWebPart.ts              # Web part entry point
│           ├── ReactPollWebPart.manifest.json   # Web part manifest
│           ├── components/
│           │   ├── ReactPoll.tsx                # Main React component
│           │   ├── ReactPoll.module.scss        # Component styles
│           │   ├── IReactPollProps.ts           # Props interface
│           │   └── Loader.tsx                   # Loading shimmer component
│           └── loc/                             # Localization strings
│               ├── en-us.js
│               └── mystrings.d.ts
├── .eslintrc.js                     # ESLint configuration
├── .gitignore                       # Git ignore rules
├── package.json                     # NPM dependencies and scripts
├── tsconfig.json                    # TypeScript configuration
└── README.md                        # This documentation
```

---

## Technology Stack

| Technology | Version | Purpose |
|------------|---------|---------|
| SharePoint Framework (SPFx) | 1.22.1 | Web part framework |
| React | 17.0.1 | UI library |
| TypeScript | 5.8.3 | Type-safe JavaScript |
| Fluent UI React | 8.x | UI components |
| PnP JS | 4.x | SharePoint REST API wrapper |
| PnP SPFx Controls | Latest | Chart control |
| Heft | Latest | Build system |
| Webpack | 5.x | Module bundler |

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 1.0.0 | January 2026 | Initial release - Upgraded to SPFx 1.22.1 |

---

## Quick Reference

### Development Commands

```bash
# Install dependencies
npm install

# Start development server
npm run start

# Build for production
npm run bundle:ship && npm run package:ship

# Clean build artifacts
npm run clean
```

### Workbench URL Template

```
https://[TENANT].sharepoint.com/sites/[SITE]/_layouts/15/workbench.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true
```

### Package Location

```
sharepoint/solution/sp-poll-webpart.sppkg
```

---

## Support

For issues and feature requests, please open an issue on GitHub.

---

## License

This project is open source and available under the MIT License.
