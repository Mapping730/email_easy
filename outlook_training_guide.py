#!/usr/bin/env python3
"""
OUTLOOK TRAINING GUIDE FOR AGENTS
==================================
This program demonstrates how to analyze a mixed Outlook setup.
It includes the actual code AND the expected results in comments
so new agents can understand the structure without running it.

TYPICAL MIXED OUTLOOK SETUP EXPLAINED:
- Multiple email account types (Exchange, IMAP, POP3, Outlook.com)
- Different folder structures per account type
- Various inbox naming conventions
- Different authentication methods
"""

import win32com.client
from datetime import datetime

def analyze_outlook_setup():
    """
    TRAINING FUNCTION: Shows how to connect and analyze Outlook accounts
    
    EXPECTED BEHAVIOR:
    This will typically find 3-5 accounts in a mixed business setup
    """
    
    print("="*70)
    print("OUTLOOK ACCOUNT ANALYSIS - TRAINING GUIDE")
    print("="*70)
    
    # STEP 1: Connect to Outlook
    print("1. Connecting to Outlook...")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        
        # EXPECTED RESULT: This usually succeeds if Outlook is installed
        print("   ✅ Connected successfully")
        
    except Exception as e:
        # COMMON ISSUE: Outlook not installed or not running
        print(f"   ❌ Connection failed: {e}")
        print("   TRAINING NOTE: This means Outlook is not available")
        return
    
    # STEP 2: Get accounts using namespace.Folders
    print("\n2. Getting accounts via namespace.Folders...")
    accounts = namespace.Folders
    print(f"   📧 Found {accounts.Count} top-level folders")
    
    """
    ACTUAL RESULTS FROM YOUR EXCHANGE SETUP (ALL ACCOUNTS ARE EXCHANGE!):
    
    Account 1: "cnc1067vm@outlook.com" (Exchange - Voicemail System)
    - Type: Exchange Server (Type 0) - NOT Outlook.com as expected!
    - Folders: 19 total (Deleted Items, Inbox, Outbox, Sent Items, Files...)
    - Inbox: 'Inbox' with 65 emails
    - Purpose: Voicemail processing system
    - Pattern: Full Exchange integration with Calendar, Contacts, Tasks
    
    Account 2: "PeteM@CNCDrywallNorth.com" (Exchange - Main Business)
    - Type: Exchange Server (Type 0)
    - Folders: 22 total (includes Calendar, Contacts, Tasks)
    - Inbox: 'Inbox' with 2,749 emails (HEAVY usage!)
    - Purpose: Main business communication
    - Pattern: Full corporate Exchange account
    
    Account 3: "Estimating2@CNCDrywallNorth.com" (Exchange - Backup Estimating)
    - Type: Exchange Server (Type 0)
    - Folders: 26 total (includes WebExtAddIns)
    - Inbox: 'Inbox' with 3,642 emails (VERY HEAVY usage!)
    - Purpose: Estimating department backup
    - Pattern: Full Exchange with custom folders
    
    Account 4: "ce2@CNCDrywallNorth.com" (Exchange - Commercial Estimating)
    - Type: Exchange Server (Type 0)
    - Folders: 29 total (includes "UD Star History From Pete")
    - Inbox: 'Inbox' with 4,729 emails (EXTREMELY HEAVY usage!)
    - Purpose: Commercial estimating system
    - Pattern: Most folders, custom organization
    
    Account 5: "Tyler Schaeffer" (Exchange - Personal Display Name)
    - Type: Exchange Server (Type 0) - NO @ SYMBOL IN NAME!
    - Folders: 25 total (Trash, Yammer Root, WebExtAddIns, Tasks, Sync Issues)
    - Inbox: 'Inbox' with 5,334 emails (HEAVIEST usage!)
    - Purpose: Individual employee account
    - Pattern: Shows as person name, not email address
    
    Account 6: "Commercial Estimator" (Exchange - Department Display Name)
    - Type: Exchange Server (Type 0) - NO @ SYMBOL IN NAME!
    - Folders: 24 total (Yammer Root, Trash, Tasks, Sync Issues)
    - Inbox: 'Inbox' with 974 emails
    - Purpose: Commercial estimating department
    - Pattern: Shows as department name, not email address
    
    🚨 CRITICAL DISCOVERY: ALL 6 ACCOUNTS ARE EXCHANGE SERVER!
    - This is a pure Exchange environment, not mixed as initially thought
    - Even cnc1067vm@outlook.com is managed through Exchange
    - namespace.Accounts only shows 4 accounts (missing Tyler & Commercial Estimator)
    - Accounts 5 & 6 appear in Folders but NOT in Accounts (Exchange display names)
    """
    
    # STEP 3: Analyze each account in detail
    print("\n3. Analyzing each account structure...")
    
    for i, account in enumerate(accounts):
        print(f"\n   Account {i+1}: {account.Name}")
        print(f"   {'='*50}")
        
        # TRAINING: Show what we're looking for
        print(f"   🔍 Analyzing folder structure...")
        
        try:
            folder_count = 0
            inbox_found = False
            inbox_name = "Not Found"
            inbox_count = 0
            folder_names = []
            
            for folder in account.Folders:
                folder_count += 1
                folder_names.append(folder.Name)
                
                # Look for inbox variations
                if folder.Name.lower() in ['inbox', 'received', 'mail']:
                    inbox_found = True
                    inbox_name = folder.Name
                    try:
                        inbox_count = folder.Items.Count
                    except:
                        inbox_count = "Access Denied"
            
            # DISPLAY ANALYSIS RESULTS
            print(f"   📁 Total folders: {folder_count}")
            print(f"   📂 Folder names: {', '.join(folder_names[:5])}{'...' if len(folder_names) > 5 else ''}")
            print(f"   📥 Inbox found: {'✅ Yes' if inbox_found else '❌ No'}")
            if inbox_found:
                print(f"   📧 Inbox name: '{inbox_name}'")
                print(f"   📊 Email count: {inbox_count}")
            
            # TRAINING: Explain what this pattern means
            # ALL ACCOUNTS ARE EXCHANGE - Update pattern detection
            if folder_count > 15:
                print(f"   🏢 PATTERN: Full Exchange Server account")
                print(f"       - Complete Outlook integration")
                print(f"       - Calendar, contacts, tasks available")
                print(f"       - Corporate email account")
                print(f"       - Heavy email usage ({inbox_count} emails)")
                
            elif 'Yammer Root' in folder_names:
                print(f"   👥 PATTERN: Exchange with Yammer integration")
                print(f"       - Social collaboration features")
                print(f"       - Modern Exchange setup")
                print(f"       - May be personal/department account")
                
            elif folder_count > 10:
                print(f"   🏢 PATTERN: Standard Exchange Server account")
                print(f"       - Full Outlook integration")
                print(f"       - Business-grade features")
                print(f"       - Moderate usage ({inbox_count} emails)")
            
        except Exception as e:
            print(f"   ❌ Error analyzing account: {e}")
            print(f"   🔒 TRAINING NOTE: This usually means access restrictions")
    
    # STEP 4: Alternative method demonstration
    print(f"\n4. Alternative: Using namespace.Accounts...")
    try:
        alt_accounts = namespace.Accounts
        print(f"   📧 Found {alt_accounts.Count} accounts via Accounts method")
        
        """
        EXPECTED DIFFERENCE:
        namespace.Folders typically returns 3-4 items (the account folders)
        namespace.Accounts typically returns 2-3 items (the actual email accounts)
        
        Sometimes namespace.Accounts gives more detailed account info:
        - Account.DisplayName: "John Smith"  
        - Account.SmtpAddress: "john@company.com"
        - Account.AccountType: 0 (Exchange), 1 (HTTP), 3 (IMAP), 4 (POP3)
        """
        
        for i, account in enumerate(alt_accounts):
            try:
                print(f"   Account {i+1}: {account.DisplayName}")
                print(f"      Email: {account.SmtpAddress}")
                print(f"      Type: {account.AccountType} ", end="")
                
                # Decode account type
                type_names = {0: "(Exchange)", 1: "(HTTP)", 3: "(IMAP)", 4: "(POP3)"}
                print(type_names.get(account.AccountType, "(Unknown)"))
                
            except Exception as e:
                print(f"   Account {i+1}: Error reading details - {e}")
        
    except Exception as e:
        print(f"   ❌ Accounts method failed: {e}")
        print(f"   📝 TRAINING NOTE: Some Outlook versions don't support this")
    
    # STEP 5: Training summary
    print(f"\n{'='*70}")
    print("TRAINING SUMMARY FOR AGENTS")
    print(f"{'='*70}")
    print("✅ WHAT YOU LEARNED:")
    print("   1. namespace.Folders gets account folders (what we usually want)")
    print("   2. Each account has different folder structures")
    print("   3. Inbox names vary: 'Inbox', 'INBOX', 'Received'")
    print("   4. Exchange accounts have more folders (Calendar, Contacts)")
    print("   5. IMAP accounts often use uppercase 'INBOX'")
    print("   6. Error handling is crucial - access can be restricted")
    print("\n🎯 KEY TAKEAWAY:")
    print("   Always check folder.Name.lower() for inbox detection")
    print("   Different account types = different behaviors")
    print("   Mixed setups are common in business environments")

def demonstrate_email_extraction():
    """
    TRAINING FUNCTION: Shows how to extract emails with expected results
    """
    print(f"\n{'='*70}")
    print("EMAIL EXTRACTION TRAINING")
    print(f"{'='*70}")
    
    """
    EXPECTED EMAIL EXTRACTION RESULTS:
    
    From Exchange account (john@company.com):
    - Email 1: "RE: Project Update" from "Sarah Manager <sarah@company.com>"
    - Email 2: "Meeting Tomorrow" from "Calendar <calendar@company.com>"  
    - Email 3: "Expense Report" from "Finance <finance@company.com>"
    
    From IMAP account (support@company.com):
    - Email 1: "Customer Issue #12345" from "customer@client.com"
    - Email 2: "System Alert" from "monitoring@server.com"
    - Email 3: "Weekly Report" from "reports@system.com"
    
    From Gmail account (personal@gmail.com):
    - Email 1: "Your Amazon Order" from "auto-confirm@amazon.com"
    - Email 2: "Newsletter" from "news@newsletter.com"
    - Email 3: "Friend's Message" from "friend@gmail.com"
    
    COMMON PATTERNS:
    - Business emails: Formal subjects, company domains
    - System emails: Auto-generated, monitoring alerts
    - Personal emails: Varied subjects, consumer services
    """
    
    print("📧 ACTUAL EMAIL PATTERNS BY ACCOUNT TYPE:")
    print("   cnc1067vm@outlook.com: Voicemail transcripts, automated system emails")
    print("   PeteM@CNCDrywallNorth.com: Main business correspondence, project emails")
    print("   Estimating2@CNCDrywallNorth.com: Backup estimating, often delivery failures")
    print("   ce2@CNCDrywallNorth.com: Supplier quotes, eQuote system responses")
    print("   Tyler Schaeffer: Personal project emails (NO @ - Exchange display name)")
    print("   Commercial Estimator: Large project updates (NO @ - Department name)")
    print("\n⚠️  AGENT ALERT: Two accounts have NO @ symbol - this is NORMAL!")
    print("   'Tyler Schaeffer' and 'Commercial Estimator' are Exchange display names")
    print("   They are NOT email addresses - they are account/department names")

def main():
    """Main training program"""
    print("🎓 OUTLOOK TRAINING PROGRAM FOR NEW AGENTS")
    print("This program shows you exactly what to expect in a mixed Outlook setup")
    print("Study the code AND the comments to understand the patterns\n")
    
    # Run the training analysis
    analyze_outlook_setup()
    demonstrate_email_extraction()
    
    print(f"\n{'='*70}")
    print("🎯 AGENT TRAINING COMPLETE")
    print("You now understand:")
    print("• How mixed Outlook setups work")
    print("• Different account types and their behaviors") 
    print("• Common folder structures and naming")
    print("• Expected email patterns per account type")
    print("• Error handling for restricted access")
    print(f"{'='*70}")

if __name__ == "__main__":
    main()