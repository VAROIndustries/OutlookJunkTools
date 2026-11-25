# OutlookJunkTools

Block, junk and unsubscribe buttons for **classic Outlook on Windows**, working the same way across every account in your profile (Exchange, Microsoft 365, IMAP, POP).

Classic Outlook keeps a separate Blocked Senders list per account, has no "block this domain" command, and has no unsubscribe button. If you run a lot of mailboxes in one profile, that gets old. This add-in adds a small **Junk Tools** group to the Home tab, the open-message tab and the right-click menu:

| Button | What it does |
|---|---|
| **Block + Junk** | Blocks the sender's address and moves the message to Junk. |
| **Block Sender** | Blocks the address. The message stays where it is. |
| **Block Domain** | Blocks everything from the sender's domain. Asks first for shared providers (gmail.com and friends) and for your own domains. |
| **Junk** | Moves the selected message(s) to that account's Junk folder. |
| **Unsubscribe** | Finds the message's unsubscribe option, shows you exactly what it will do, and waits for you to confirm. |
| **Block List** | View, remove or manually add blocked addresses and domains. |

Mail from a blocked sender is moved to Junk as it arrives, in any account. Unread mail from the last 7 days is also checked shortly after Outlook starts.

## Status

Early release (0.2.x). The rules engine has unit tests. In real use so far: the add-in loads, and all three unsubscribe routes (one-click, email, browser) have worked against live mailing lists. The block and junk commands and the incoming-mail check have had little real-world use yet. Expect rough edges and please open an issue.

## How blocking works

The add-in keeps **its own block list** (`%AppData%\OutlookJunkTools\blocklist.json`). Outlook's object model does not expose its built-in Blocked Senders lists, so the add-in enforces its list itself. That means:

- Blocking only acts while Outlook is running on this PC.
- Blocked mail is moved to Junk after it is delivered. Nothing is rejected at the server.
- The list does not sync between computers.
- You cannot block one of your own account addresses.

## How unsubscribe works

When you press Unsubscribe on a single message, the add-in looks for, in order:

1. **One-click** ([RFC 8058](https://www.rfc-editor.org/rfc/rfc8058)): an HTTPS `List-Unsubscribe` link plus `List-Unsubscribe-Post`. The add-in sends the one-click POST for you.
2. **Email**: a `mailto:` in `List-Unsubscribe`. The add-in sends the unsubscribe email from the account that received the message.
3. **Header link**: a web link in `List-Unsubscribe`. Opened in your browser.
4. **Body link**: the first link in the message whose text or address mentions unsubscribe, opt out or manage preferences. Opened in your browser.

Before anything happens you see the sender, the method, the full destination and whether your mail server recorded a DKIM pass. If one method fails, the next one is offered.

Safety notes:

- Nothing is ever sent or opened without a button press and a confirmation, and never in bulk.
- The only network request the add-in makes itself is the confirmed one-click POST: HTTPS only, no cookies, no credentials, no redirects followed, 15 second timeout.
- The DKIM line repeats what your mail server wrote in `Authentication-Results`. The add-in does not verify signatures itself.
- Unsubscribing tells the sender your address is live. For obvious spam, use **Block + Junk** instead.

No telemetry. The block list and a log (`%AppData%\OutlookJunkTools\log.txt`) stay on your PC.

## Requirements

- Classic Outlook for Windows (Microsoft 365 / Office 2016 or later), 64-bit or 32-bit. **Not** the new Outlook and not Outlook on the web; neither can load COM add-ins.
- .NET Framework 4.8 (included with Windows 10 1903 and later, and Windows 11).
- To build: the [.NET SDK](https://dotnet.microsoft.com/download). No Visual Studio and no Office developer workload needed.

## Build and install

```powershell
git clone https://github.com/VAROIndustries/OutlookJunkTools.git
cd OutlookJunkTools
dotnet build -c Release
```

Close Outlook, then from a normal PowerShell window:

```powershell
powershell -ExecutionPolicy Bypass -File .\install.ps1
```

Start Outlook. `install.ps1` copies the build to `%LocalAppData%\OutlookJunkTools\app` and registers the add-in for the current user only (HKCU). No admin rights.

To update, pull, rebuild, close Outlook and run `install.ps1` again. To remove, close Outlook and run `uninstall.ps1`. Your block list and log are left in place.

> Run the install script from an ordinary PowerShell window. Run from inside a packaged (MSIX / Store) app's shell, Windows redirects the registry and AppData writes into that app's private sandbox and Outlook never sees them.

## Troubleshooting

- **No Junk Tools group.** Check File > Options > Add-ins. If *Outlook Junk Tools* is not listed at all, the install did not reach your real registry (see the note above). If it is listed as Inactive or Disabled, enable it from Manage: COM Add-ins > Go.
- **Something failed.** Look at `%AppData%\OutlookJunkTools\log.txt`.
- **Ribbon errors.** File > Options > Advanced > Developers > "Show add-in user interface errors" makes Outlook report ribbon problems.

## How it is built

- `src/OutlookJunkTools.Core`: block matching, list storage, guard rules, unsubscribe header parsing. No Outlook dependency.
- `tests/OutlookJunkTools.Core.Tests`: xUnit tests for Core. Run with `dotnet test`.
- `src/OutlookJunkTools.AddIn`: a plain COM add-in (`IDTExtensibility2` + ribbon XML). No VSTO.

## License

MIT. Not affiliated with or endorsed by Microsoft.
