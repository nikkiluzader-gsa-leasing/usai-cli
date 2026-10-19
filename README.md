# USAi CLI

Chat with [USAi](https://api.gsa.usai.gov) models from PowerShell. Replies stream in as they're written, you can attach files with `@`, and the model can (with your approval) read and edit the files in your current folder and run `git` and `clasp` commands for you.

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
| `/export [full] [path]` | Save the conversation as Markdown, with one line per tool step and how it went. `full` adds each step's raw input, which can include file contents |
| `/context` | Show conversation size, settings, and anything waiting to be sent |
| `/system [text \| reset]` | Show or change the system prompt |
| `/temperature <0-1 \| default>` | Set response creativity |
| `/max-tokens <n \| default>` | Limit reply length. A reply that hits the limit is continued automatically (up to twice) |
| `/stream on\|off` | Turn streaming on or off |
| `/tools on\|read\|off` | Let the model use files and git/clasp in the current folder; `read` means look only (see below) |
| `/approve [ask\|reads\|edits\|all]` | Choose what runs without asking (see below). On its own, pick with the arrow keys |
| `/revert` | Undo the last file change or delete the model made (repeat to go further back) |
| `/permissions [remove <n> \| reset]` | Show or remove your "always" approvals |
| `/clasp [filter]` | Pick one of your Apps Script projects with the arrow keys, then clone it into its own folder (or pull it, if it's already here). The model is told which project you picked |
| `/compact [focus]` | Replace a long conversation with a summary, to save budget and keep the model focused |
| `/init` | Have the model write a `USAI.md` with instructions for this folder (see below) |
| `/cd`, `/ls`, `/pwd` | Move around and look at folders without leaving the chat |
| `/key reset` | Replace your saved API key |
| `/help` | Show all commands |
| `/exit` | Quit (also `/quit`) |

## Tools

When tools are on (the default), the model can work with the files in your current folder, and with `git` and `clasp` if they're installed:

| Tool | What it does | How you approve it |
| --- | --- | --- |
| **list_files**, **read_file**, **search_files** | Look at files | **y** or Enter |
| **edit_file**, **write_file**, **delete_file** | Change, create or delete a file | Shows a red/green diff (or the file being deleted) first, then **y**. Enter alone doesn't count |
| **run_command** | Run an allowed `git` or `clasp` subcommand | **y** for each command, or Enter for read-only ones like `git status`. `git push`, `clasp push` and deployments make you type `push` or `deploy` |
| **ask_user** | Ask you to choose, e.g. which Apps Script project | Pick with the arrow keys, or choose "Type something else" |

Instead of **y** you can also answer:

- **a**: always allow this for the rest of the session
- **f**: always allow this in the current folder and its subfolders. This one is saved, so it still applies next time you start `usai` here
- **n**: no. You can then type what the model should do instead (or just press Enter), and it's passed on

What "this" covers depends on the step: one read tool (`read_file`), **all file edits** (deletes included), or one command (`git commit`, which also covers `git add`). The diff is still shown for approved edits, and `/revert` still works. `git push`, `clasp push` and deployments can't be pre-approved; they ask for the word every time. **f** isn't offered in your home folder or at the top of a drive, where it would cover too much. `/permissions` lists what you've allowed and removes it.

To stop being asked about whole kinds of steps, set a mode with `/approve` (it's saved, and shown before the `>` prompt):

| Mode | Runs without asking |
| --- | --- |
| `ask` (default) | Nothing |
| `reads` | Reading files and read-only commands like `git status` and `clasp list` |
| `edits` | Also file changes. You still see each diff, and `/revert` undoes them |
| `all` | Also commands like `git commit` and `clasp pull` |

`git push`, `clasp push` and deployments ask for the word in every mode.

So you can ask for a whole job, and approve each step as it goes:

```
> pull my Apps Script project, fix the date bug in Code.js, and push it
  * run_command clasp pull                         Run this command? [y]es [a] [f] [n]o
  * read_file Code.js                              Allow? [y]es [a] [f] [n]o
  * edit_file Code.js        (diff shown)          Save these changes? [y]es [a] [f] [n]o
  * run_command clasp push                         Type push to confirm
```

Limits, enforced by the script rather than by asking the model nicely:

- Everything stays inside the current folder and its subfolders (use `/cd` to change it).
- Credential files are never read or changed: `.clasprc.json` (your Google sign-in), `.env` files, `credentials.json`, keys and certificates, and USAi CLI's own `~\.usai` folder. Nothing inside `.git`, and not `.clasp.json`, can be edited either.
- Only `git` and `clasp` can run, and only allowlisted subcommands. Commands start directly, never through a shell, so the model can't chain or redirect anything.
- Commands run in the current folder, or in a subfolder the model names (say, a clasp project cloned into `testing\`). Before `clasp push` or a deployment, the confirmation shows the script ID from the `.clasp.json` that clasp will use, so you can check it's the project you mean.
- Blocked: `git reset`, `clean`, `checkout`, `restore`, `rm`, `rebase`, `config`, force pushes and branch deletion; `clasp run`, `login`, `logout`, `create`, `delete` and `--force`. The model is told why, so it can tell you the command to run yourself.
- Commands that wait for typed input fail instead of hanging. Esc stops a running command, and any command is stopped after the request timeout (300 seconds).

Before every file change or delete, the original is copied to `~\.usai\backups`. `/revert` undoes the model's last change in this session (bringing back a deleted file); run it again to go further back. To change a long block, the model can give a range of line numbers instead of copying the old text, but only right after reading the file, so the numbers can't be stale. `write_file`, which replaces a whole file, only works on a file the model has read in this conversation and that hasn't changed since; otherwise it's told to read the file first.

A long job can take many steps. After 50 for one message, `usai` asks **Keep going?** so you can let it carry on without typing "continue".

`/tools read` turns off editing and commands but keeps the three look-only tools (and **ask_user**).

## Project instructions (USAI.md)

Put a `USAI.md` file in a folder, and its contents are added to the model's instructions in every conversation started there, or in any folder below it. Use it for what you'd otherwise repeat each time: what the project is, which files matter, how you push, and style rules. `~\.usai\USAI.md` applies everywhere. When several apply, the more specific one comes last. The banner shows which ones are loaded.

`/init` has the model look around the folder and write a `USAI.md` for you, with the usual diff and approval. A `USAI.md` that came with someone else's project is instructions you didn't write, so read it before trusting it. Every step still needs your approval either way. Models that don't support tool calling are detected automatically and just chat normally. Print mode (`-p`) never uses tools, since nobody is there to approve them.

## What gets sent where

The script only talks to `https://api.gsa.usai.gov`, using two endpoints:

- `GET /api/v1/models` for the model list
- `POST /api/v1/chat/completions` for everything else

Each request includes the conversation so far, plus any files you attached, any file contents you approved the model reading, and the output of commands you approved. Nothing else leaves your machine, except what the `git` and `clasp` commands you approve send themselves (for example, `git push` to your git server and `clasp push` to Google).

Stored locally in `~\.usai`:

| File | Contents |
| --- | --- |
| `api-key.xml` | Your API key, encrypted with Windows DPAPI (only your Windows account on this machine can read it) |
| `config.json` | Your settings: default model, system prompt, temperature, folders where you chose **f**, and so on |
| `USAI.md` | Optional instructions for every conversation (see above) |
| `sessions\*.json` | Saved conversations (images and PDFs are left out) |
| `backups\*` | Copies of files from before the model changed them |
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
| An update or a change to `usai.ps1` seems to have no effect | Typing `usai` runs the installed copy in `~\.usai\bin`, not the file in the folder you're in. Run `.\usai.ps1 -Install` again after updating, or run `.\usai.ps1` directly to try a local copy. The banner warns when the folder's `usai.ps1` differs from the one running. |
| Garbled characters | Use Windows Terminal, or a font with good Unicode coverage. |
| Blocked as malicious content by antivirus | Let your IT or security team know; don't add an exclusion yourself. |

## How it works

USAi's chat API follows the OpenAI Chat Completions format. The script:

- streams replies as server-sent events, and falls back to a single JSON response when a model doesn't stream
- uses the API's tool calling for the tools, and sends back any provider-specific data a model includes with a tool request (Gemini models need this)
- spaces requests out to stay under the rate limit, and retries brief outages (429, 503) with a short wait
- uses `System.Net.Http.HttpClient` with explicit UTF-8, so accented letters and curly quotes come through correctly on Windows PowerShell 5.1

## License

This project is in the public domain within the United States, and copyright and related rights in the work worldwide are waived through the [CC0 1.0 Universal public domain dedication](https://creativecommons.org/publicdomain/zero/1.0/). See [LICENSE.md](LICENSE.md).
