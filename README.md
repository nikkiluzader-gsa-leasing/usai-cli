# USAi CLI

Chat with [USAi](https://api.gsa.usai.gov) models from PowerShell. Replies stream in as they're written, you can attach files with `@`, and the model can (with your approval) read the files in your current folder to answer questions about them.

It's a single PowerShell script with no dependencies. It works in Windows PowerShell 5.1 and PowerShell 7+.

New to it? Start with the **[User Guide](USER_GUIDE.md)**. This README is the quick reference.

```
> why does @build.ps1 fail when the output folder already exists?

The script calls New-Item without -Force on line 14, so it throws if
the folder is there...
  claude_4_6_sonnet  |  1,204 in / 312 out tokens  |  4.8s
```

## Install

You need a USAi API key from the USAi Console (**Keys > My API keys**).

1. Download `usai.ps1`.
2. In PowerShell, go to the folder you saved it in and run:

   ```powershell
   .\usai.ps1 -Install
   ```

3. Open a new PowerShell window and type `usai`. The first time, it asks for your API key.

`-Install` copies the script to `~\.usai\bin` and adds that folder to your user PATH, so `usai` works from any folder. To update, run a newer copy with `-Install` again. To remove it, run `usai -Uninstall`.

**If Windows won't run the script:** files downloaded from the internet or email are marked as untrusted, and many machines block marked scripts. Right-click `usai.ps1`, choose **Properties**, tick **Unblock**, click **OK**, and try again. If it's still blocked, your execution policy is set by your organization; ask your IT team rather than working around it.

**If you can't change your PATH:** run `.\usai.ps1 -Install` anyway (the copy step still works), then add this line to your PowerShell profile (`notepad $PROFILE`):

```powershell
function usai { & "$HOME\.usai\bin\usai.ps1" @args }
```

## Usage

```powershell
usai                                   # start chatting
usai "explain @.\script.ps1"           # start with a first message
usai -c                                # continue your most recent conversation
usai -m claude_4_6_sonnet              # use a different model for this session
usai -p "summarize this week's notes"  # print the answer and exit (for scripts)
Get-Content .\error.log | usai -p "what went wrong?"
```

### While chatting

| Input | What it does |
| --- | --- |
| `@path` | Attach a file to your message, e.g. `summarize @notes.txt`, `@"my file.pdf"`, `@src\*.ps1` |
| `"""` | Start or end a multi-line message (pasting several lines also works) |
| `\` at end of line | Continue your message on the next line |
| `Esc` | Stop a reply that's in progress (Ctrl+C also works while a reply runs) |
| Up arrow | Recall what you typed earlier |

Attachments can be text files, images (png, jpg, gif, webp) or PDFs. Whether a model can read images and PDFs depends on the model.

### Commands

| Command | What it does |
| --- | --- |
| `/model` | Pick a model with the arrow keys; type to filter. `/model sonnet` switches directly if only one model matches |
| `/models` | List available models |
| `/clear` | Start a new conversation (also `/new`) |
| `/resume` | Pick a saved conversation to continue |
| `/retry` | Resend a failed message, or regenerate the last reply |
| `/undo` | Remove the last exchange |
| `/file <paths>` | Attach files to your next message (also `/add`; wildcards work) |
| `/drop` | Discard attachments you haven't sent yet |
| `/copy` | Copy the last reply to the clipboard |
| `/export [path]` | Save the conversation as Markdown |
| `/context` | Show conversation size, settings, and anything waiting to be sent |
| `/system [text \| reset]` | Show or change the system prompt |
| `/temperature <0-1 \| default>` | Set response creativity |
| `/max-tokens <n \| default>` | Limit reply length |
| `/stream on\|off` | Turn streaming on or off |
| `/tools on\|off` | Let the model read files in the current folder (see below) |
| `/approve on\|off` | Skip the approval prompt for those reads |
| `/cd`, `/ls`, `/pwd` | Move around and look at folders without leaving the chat |
| `/key reset` | Replace your saved API key |
| `/help` | Show all commands |
| `/exit` | Quit (also `/quit`) |

## File tools

When file tools are on (the default), the model can use three read-only tools to answer questions about your files:

- **list_files**: list a folder's contents
- **read_file**: read a text file
- **search_files**: search file contents

Every request shows what the model wants to read and waits for you to answer **y** (yes), **a** (always allow this tool for the rest of the session), or **n** (no). The tools:

- only work inside the current folder and its subfolders (use `/cd` to change it)
- can't change files or run commands
- skip folders like `.git`, `node_modules`, `bin` and `obj` when searching

Models that don't support tool calling are detected automatically and just chat normally. Print mode (`-p`) never uses tools, since nobody is there to approve them.

## What gets sent where

The script only talks to `https://api.gsa.usai.gov`, using two endpoints:

- `GET /api/v1/models` for the model list
- `POST /api/v1/chat/completions` for everything else

Each request includes the conversation so far, plus any files you attached and any file contents you approved the model reading. Nothing else leaves your machine.

Stored locally in `~\.usai`:

| File | Contents |
| --- | --- |
| `api-key.xml` | Your API key, encrypted with Windows DPAPI (only your Windows account on this machine can read it) |
| `config.json` | Your settings: default model, system prompt, temperature, and so on |
| `sessions\*.json` | Saved conversations (images and PDFs are left out) |
| `bin\usai.ps1` | The installed copy of the script |

You can also set the key with the `USAI_API_KEY` environment variable, which takes priority over the saved key.

**Don't send sensitive, classified, or otherwise restricted information.** Follow your agency's rules for what can go to USAi.

## Troubleshooting

| Problem | Fix |
| --- | --- |
| `401` or `403` errors | USAi keys expire every 7 days. Rotate yours in the USAi Console, then run `/key reset`. |
| `429` errors | You hit the rate limit (3 calls per second) or your weekly budget. The script retries rate limits on its own; a budget limit resets on Monday. |
| `503` errors | That model is temporarily unavailable. The script retries twice; after that, try `/retry` later or switch with `/model`. |
| A model rejects temperature, max tokens, streaming or tools | The script notices, drops that option for that model, and resends. `/context` shows what was dropped. |
| Garbled characters | Use Windows Terminal, or a font with good Unicode coverage. |
| Blocked as malicious content by antivirus | Let your IT or security team know; don't add an exclusion yourself. |

## How it works

USAi's chat API follows the OpenAI Chat Completions format. The script:

- streams replies as server-sent events, and falls back to a single JSON response when a model doesn't stream
- uses the API's tool calling for the file tools, and sends back any provider-specific data a model includes with a tool request (Gemini models need this)
- spaces requests out to stay under the rate limit, and retries brief outages (429, 503) with a short wait
- uses `System.Net.Http.HttpClient` with explicit UTF-8, so accented letters and curly quotes come through correctly on Windows PowerShell 5.1

## License

This project is in the public domain within the United States, and copyright and related rights in the work worldwide are waived through the [CC0 1.0 Universal public domain dedication](https://creativecommons.org/publicdomain/zero/1.0/). See [LICENSE.md](LICENSE.md).
