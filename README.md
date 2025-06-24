# 📧 Exchange Online Journaling & Reporting Suite

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Exchange Online](https://img.shields.io/badge/Exchange%20Online-Supported-green.svg)](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/yourusername/exchange-journaling/graphs/commit-activity)

A comprehensive PowerShell solution for configuring Exchange Online journaling and generating automated monthly HTML reports with detailed message analytics and user insights.

## 🌟 Features

### 📋 Journaling Configuration
- ✅ **Global Journaling Rules** - Capture all incoming and outgoing messages
- ✅ **Automatic Prerequisites** - Handles undeliverable journal reports configuration
- ✅ **Mailbox Validation** - Ensures all required mailboxes exist before configuration
- ✅ **Interactive Setup** - Prompts for missing information during configuration
- ✅ **Rule Management** - Creates and updates journaling rules seamlessly

### 📊 Advanced Monthly Reporting
- ✅ **Full Monthly Coverage** - Always generates complete calendar month reports
- ✅ **Enhanced User Analysis** - Detailed information about top senders and recipients
- ✅ **Professional HTML Reports** - Modern, responsive design with interactive elements
- ✅ **Message Statistics** - Comprehensive breakdown of inbound/outbound communications
- ✅ **Activity Metrics** - Daily averages, unique contacts, and communication patterns
- ✅ **External/Internal Identification** - Visual indicators for external vs internal users
- ✅ **Historical Data Support** - Access to data beyond the standard 10-day limit

### 🤖 Automation Features
- ✅ **Scheduled Tasks** - Automated monthly report generation
- ✅ **Multiple Creation Methods** - XML, PowerShell cmdlets, and schtasks.exe fallbacks
- ✅ **Error Handling** - Robust error management and recovery
- ✅ **Task Verification** - Automatic testing of created scheduled tasks

## 📸 Sample Report Features

### Enhanced User Analysis
- **Display Names & Titles** - Full user information including department and office
- **Message Breakdowns** - Detailed inbound/outbound statistics
- **Activity Patterns** - Daily averages and unique contact counts
- **External Indicators** - Clear identification of external communications

### Professional Dashboard
- **Modern Design** - Clean, gradient-based interface
- **Interactive Charts** - Visual representation of daily message volumes
- **Responsive Layout** - Works on desktop and mobile devices
- **Comprehensive Statistics** - Six key metrics at a glance

## 🚀 Quick Start

### Prerequisites

- **PowerShell 5.1** or later
- **Exchange Online Management Module**
- **Exchange Online Administrator** permissions
- **Compliance Administrator** permissions (for journaling)
- **Windows 10/11** or **Windows Server 2016+**

### Basic Usage

```powershell
# Configure journaling and generate current month report
.\Configure-ExchangeJournaling.ps1 -JournalEmailAddress "journal@yourdomain.com" -UndeliverableReportsAddress "undeliverable@yourdomain.com"

# Generate reports only (skip journaling configuration)
.\Configure-ExchangeJournaling.ps1 -JournalEmailAddress "journal@yourdomain.com" -SkipJournalingConfig

# Set up automated monthly reports
.\Schedule-JournalingReports.ps1 -JournalEmailAddress "journal@yourdomain.com"
