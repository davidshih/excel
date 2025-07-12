# SharePoint Integration Guide

## Overview

The enhanced Excel splitter now includes direct SharePoint integration, eliminating the need for PowerShell scripts. The new features include:

- Browser-based authentication using Microsoft SSO
- Automatic email lookup from Microsoft 365
- Interactive reviewer selection with checkboxes
- Direct API calls to share folders
- Email notifications to reviewers
- Real-time progress tracking

## Setup Requirements

### 1. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to Azure Active Directory > App registrations
3. Click "New registration"
4. Configure:
   - Name: `Excel Splitter SharePoint Integration`
   - Supported account types: Single tenant
   - Redirect URI: `http://localhost:8400` (Web platform)

### 2. API Permissions

Add the following Microsoft Graph API permissions:
- `User.Read` (Delegated)
- `User.ReadBasic.All` (Delegated)
- `Sites.ReadWrite.All` (Delegated)
- `Mail.Send` (Delegated)

Grant admin consent for your organization.

### 3. Authentication Settings

Enable the following in Authentication settings:
- Access tokens
- ID tokens
- Live SDK support (optional)

## Using the Enhanced Notebook

### File: `excel_splitter_interface_sharepoint_enhanced.ipynb`

### Step 1: Configure Azure AD
```
Client ID: [Your app's client ID]
Tenant ID: [Your Azure AD tenant ID]
```

### Step 2: Authenticate
- Click "Authenticate" button
- Browser window opens for Microsoft login
- Sign in with your Microsoft 365 account
- Grant permissions if prompted

### Step 3: Configure SharePoint
```
SharePoint Site: https://yourcompany.sharepoint.com/sites/yoursite
```

### Step 4: Process Excel File
1. Select your Excel file
2. Specify the reviewer column (default: "Reviewer")
3. Click "Split Excel File"

### Step 5: Share Folders
After processing, you'll see:
- List of all reviewers with checkboxes
- Auto-populated email addresses (where found)
- Manual email entry for missing addresses
- "Share to Selected Reviewers" button

### Step 6: Monitor Progress
- Real-time progress bar
- Status updates for each reviewer
- Email notification confirmation

## Features in Detail

### Email Lookup
- Searches Microsoft 365 directory by display name
- Handles partial matches and duplicates
- Falls back to manual entry when not found

### Folder Sharing
- Uses SharePoint REST API
- Sets "Edit" permissions by default
- Creates sharing invitations
- Maintains audit trail

### Email Notifications
- Professional HTML email template
- Includes direct link to SharePoint folder
- Explains co-authoring synchronization
- Sent from authenticated user's account

### Error Handling
- Graceful fallback for missing emails
- Retry logic for API calls
- Detailed error messages
- Continue processing on individual failures

## Comparison: Old vs New

### Old Method (PowerShell Script)
- Generated script file
- Required PowerShell modules
- Manual execution needed
- Separate email lookup process
- No progress tracking

### New Method (Direct API)
- Integrated in Jupyter notebook
- Uses browser SSO
- One-click sharing
- Automatic email lookup
- Real-time progress

## Troubleshooting

### Authentication Issues
- Ensure Client ID and Tenant ID are correct
- Check API permissions are granted
- Try clearing browser cache
- Use InPrivate/Incognito mode

### Email Lookup Failures
- Verify user exists in Microsoft 365
- Check for typos in display names
- Try different name formats
- Use manual email entry

### Sharing Failures
- Confirm SharePoint site URL is correct
- Verify you have sharing permissions
- Check folder names for special characters
- Ensure folders were created successfully

### Rate Limiting
- Built-in delays between API calls
- Process in smaller batches if needed
- Wait and retry if limits hit

## Security Considerations

- Uses OAuth 2.0 for authentication
- Tokens stored only in memory
- No credentials saved to disk
- Follows Microsoft security best practices
- Respects SharePoint permission model

## Next Steps

1. Test with a small group first
2. Monitor email delivery
3. Verify folder access works
4. Collect reviewer feedback
5. Adjust email template if needed