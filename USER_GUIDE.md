# USAi CLI User Guide

This guide walks you through using USAi CLI day to day. You don't need to know PowerShell beyond opening a window and typing a command.

**Contents**

1. [What it is](#1-what-it-is)
2. [Set it up](#2-set-it-up)
3. [Your first conversation](#3-your-first-conversation)
4. [Choosing a model](#4-choosing-a-model)
5. [Working with your files](#5-working-with-your-files)
6. [Writing longer messages](#6-writing-longer-messages)
7. [Managing conversations](#7-managing-conversations)
8. [Adjusting how it responds](#8-adjusting-how-it-responds)
9. [Using it in scripts](#9-using-it-in-scripts)
10. [Keeping your API key current](#10-keeping-your-api-key-current)
11. [Tips for better answers](#11-tips-for-better-answers)
12. [Troubleshooting](#12-troubleshooting)
13. [Quick reference](#13-quick-reference)

---

## 1. What it is

USAi CLI lets you chat with the AI models available through USAi from a PowerShell window instead of a web browser. It's especially handy when you're working with files: you can hand it a document or script, or let it look through a folder to answer questions about what's there.

What you type goes to USAi, along with any files you attach or allow it to read, and the answer streams back into your window. Nothing goes anywhere else.

> **Before you start:** only send information you're allowed to share with USAi. Don't send sensitive, classified, or otherwise restricted information.

## 2. Set it up

You only do this once.

**Get an API key.** Sign in to the USAi Console and go to **Keys > My API keys** to create one. Keep the page open; you'll paste the key in a moment.

**Install the tool.**

1. Save `usai.ps1` somewhere, such as your Downloads folder.
2. Open PowerShell. (Press the Windows key, type `PowerShell`, and press Enter.)
3. Go to the folder where you saved it. For Downloads, type:

   ```powershell
   cd ~\Downloads
   ```

4. Run the installer:

   ```powershell
   .\usai.ps1 -Install
   ```

5. Close PowerShell and open a new window.
6. Type `usai` and press Enter.
7. Paste your API key when asked (right-click pastes in most PowerShell windows), then press Enter. It's saved, encrypted, so you won't be asked again until you replace it.

**Updating later:** type `usai -Update` in PowerShell, or `/update` while you're chatting. `usai` also tells you, at most once a day, when a newer version is available. (Running `usai -Install` again doesn't update; it just reinstalls the copy you already have.)

**If step 4 says the script "cannot be loaded" or "is not digitally signed":** Windows has marked the downloaded file as untrusted. In File Explorer, right-click `usai.ps1`, choose **Properties**, tick **Unblock** near the bottom, and click **OK**. Then try step 4 again. If it still won't run, ask your IT team; don't try to work around the block.

## 3. Your first conversation

Type `usai` in any PowerShell window. You'll see something like:

```
  USAi  chat in your terminal
  model: gpt_5_5_default_v2
  cwd:   C:\Users\you\Documents
  tools: can read files here, asks first

  /help for commands  -  @file to attach  -  Esc to interrupt

>
```

Type a question after the `>` and press Enter. The answer appears as it's written. Under each answer is a gray line showing which model answered, how many tokens it used, and how long it took.

Keep going: each message you send includes the earlier ones, so you can ask follow-ups like "make that shorter" or "now do the same for the second table."

**Useful keys**

- **Esc** stops an answer partway through, if it's going in the wrong direction.
- **Up arrow** brings back what you typed before, so you can edit and resend it.
- Type **/exit** to quit. Your conversation is saved automatically.

## 4. Choosing a model

USAi offers models from several providers (OpenAI, Anthropic, Google, Meta and xAI), and they have different strengths. To switch:

- Type **/model**. A list appears; use the **arrow keys** to move, start **typing to filter** it, then press **Enter** to choose or **Esc** to cancel.
- Or type part of a name: **/model sonnet** or **/model gemini pro**. If only one model matches, it switches right away; otherwise the list opens, filtered to those models.

Your choice is remembered next time. To use a different model just once, start with `usai -m <model name>`.

> Not sure which to pick? Try the same question on two models and compare. Some write better, some reason better, and some are faster.

## 5. Working with your files

There are two ways to give the model your files.

### Attach a file to a message

Put `@` in front of a file name anywhere in your message:

```
> summarize @meeting-notes.txt in five bullet points
> compare @draft-v1.txt and @draft-v2.txt
> what's in this screenshot? @error.png
```

- If the name has spaces, use quotes: `@"Q3 report.pdf"`
- You can attach several files with a wildcard: `@reports\*.txt`
- Files are looked up starting from your current folder (shown on the `cwd:` line). Use `/cd` to change folders, `/ls` to see what's there, and `/pwd` to check where you are.

You can attach plain text files (including code, CSV and Markdown), images (PNG, JPG, GIF, WebP) and PDFs. Not every model can read images or PDFs; if one can't, try another. Word and Excel files can't be read directly; save them as PDF or text first.

If you'd rather add files before typing your question, use `/file`:

```
> /file budget.csv notes.txt
  + budget.csv
  + notes.txt
> which line items went over budget?
```

`/drop` discards files you've added but haven't sent yet.

### Let it look around a folder

You can also just ask about your files, and the model will look for what it needs:

```
> which of these scripts send email?
  * search_files "Send-MailMessage"
    Allow? [y]es  [a] search_files this session  [f] search_files in this folder  [n]o
```

Each time it wants to look at something, it tells you what and waits for your answer:

- **y** (or Enter): allow this one request.
- **a**: allow this kind of request for the rest of this session.
- **f**: allow this kind of request in this folder from now on, including in future sessions.
- **n**: refuse. You can then type what it should do instead (or just press Enter), and the model follows that.

It can **list** folders, **read** text files and **search** file contents. It can only see the current folder and the folders inside it, and it never reads files that hold passwords or sign-ins, such as `.clasprc.json` or `.env`.

### Let it change files

Ask for a change, and the model can make it. Before anything is saved, you see what would change, removed lines in red and added lines in green:

```
> the onEdit trigger in Code.js logs twice, fix it
  * edit_file Code.js
      1   function onEdit(e) {
      2 -   log(e);
      3     log(e);
      4   }
    Save these changes to Code.js? [y]es  [n]o
```

Press **y** to save or **n** to refuse. Enter alone doesn't count, so you can't save a change by accident.

Tired of saying yes? Two more answers save you from being asked again:

- **a**: allow **all file edits** for the rest of this session.
- **f**: allow all file edits in **this folder** (and the folders inside it), now and every time you use `usai` here. It's remembered until you remove it.

You still see every diff, so you can catch anything odd. `/revert` still works too.

To stop the questions about edits for good, type `/approve edits` (see [Fewer questions](#fewer-questions) below).

The model can also **delete** a file when you ask it to. That's asked like an edit, and `/revert` brings the file back.

Changed your mind afterwards? Type `/revert` to put the file back the way it was. Each `/revert` undoes one more change from this session. A copy of each original is also kept in `~\.usai\backups`.

### Let it run git and clasp

If you have `git` or Google's `clasp` installed, the model can run them for you, so you can ask for a whole job in one message:

```
> pull my Apps Script project, fix the date bug in Code.js, and push it
```

It runs `clasp pull`, reads the files, shows you the fix, and runs `clasp push`, asking you at each step:

- Commands that only look, like `git status` or `clasp status`, ask like file reads; Enter means yes.
- Commands that change something, like `git commit` or `clasp pull`, need **y**. Here too, **a** allows that one command (say, every `git commit`) for the session, and **f** allows it in this folder for good. Allowing `git commit` also allows `git add`, since a commit needs it, but no other commands.
- `git push`, `clasp push`, and creating versions or deployments make you **type** `push` or `deploy`. `clasp push` replaces the live code that your sheets and forms run, so take a moment to check before you type it. These can't be pre-approved with **a** or **f**.

### Pick an Apps Script project

Type `/clasp` to see your Apps Script projects and pick one with the arrow keys (type part of a name to narrow the list). Then pick what should happen:

- **Clone into a new folder**: clasp copies the project into a folder named after it, inside the current folder.
- **Clone into this folder**: offered when the current folder isn't already a clasp project.
- **Pull the latest code** or **Use it as it is**: offered when you've already cloned that project here.

The model is told which project you picked and where it is, so you can go straight to "fix the date bug in Code.js". `/clasp docu` starts the list filtered to names containing "docu".

If you'd rather ask in words ("which scripts do I have?"), the model lists them and then lets you pick with the arrow keys too. It does this whenever you need to choose between options. You can always choose **Type something else** and answer in your own words.

Commands that could throw away work, like `git reset` or `git checkout`, are blocked, as are `clasp run`, `clasp login` and any program other than git and clasp. When something is blocked, the model tells you the command so you can run it yourself if you really want to.

### Turning tools down or off

- `/tools read` lets the model look at files but not change them or run commands.
- `/tools off` turns tools off completely; `/tools on` turns everything back on.

### Fewer questions

`/approve` sets what the model can do without asking. Type it on its own to pick with the arrow keys, or name the mode:

- `/approve ask`: ask before every step (the default).
- `/approve reads`: reading files and look-only commands don't ask.
- `/approve edits`: file changes don't ask either. You still see every diff, and `/revert` still works.
- `/approve all`: commands like `git commit` and `clasp pull` don't ask either.

`git push`, `clasp push` and deployments always ask you to type the word, in every mode. The mode is remembered, and while it's not `ask` it's shown before the `>` prompt (for example `edits >`), so you always know.

On a long job the model may take many steps. After 50 steps for one message it asks **Keep going?** Press **y** to let it carry on.

`/permissions` lists everything you've answered **a** or **f** to. `/permissions remove 2` forgets one saved folder, and `/permissions reset` forgets everything.

### Give it standing instructions

If you find yourself telling the model the same things each time ("this is the leasing tracker script, Code.js has the triggers, always run clasp status before pushing"), put them in a file named `USAI.md` in that folder. Everything in it is added to every conversation you start there or in a folder inside it. A `USAI.md` in `~\.usai` applies everywhere.

Type `/init` and the model will look around and write a first version for you to approve and edit.

> Everything the model reads, and the output of every command it runs, is sent to USAi, just like an attachment. Be careful with **a** and `/approve` in folders that contain information you shouldn't share.

## 6. Writing longer messages

Pressing Enter sends your message. For anything longer than one line:

- **Paste it.** Pasting several lines at once sends them together as one message.
- **Use triple quotes.** Type `"""` on its own line, write as many lines as you like, then type `"""` again to send:

  ```
  > """
  . Rewrite this email to sound friendlier:
  .
  . Per my last message, the report is late.
  . """
  ```

- **End a line with `\`** to keep typing on the next line.

## 7. Managing conversations

Every conversation is saved automatically as you go.

| To... | Type |
| --- | --- |
| Start fresh | `/clear` |
| Go back to an earlier conversation | `/resume`, then choose a number |
| Pick up your last conversation when starting | `usai -c` |
| Try the last answer again | `/retry` |
| Take back your last question and its answer | `/undo` |
| Copy the last answer to the clipboard | `/copy` |
| Save the conversation as a document | `/export` (or `/export notes.md`). Each step the model took is one line, like `edit_file Code.js -> Saved`. `/export full` also includes what each step sent, which can include file contents, so don't share that version |
| See how long the conversation is and your settings | `/context` |
| Shrink a long conversation into a summary and keep going | `/compact` (or `/compact focus on the Code.js changes`) |

**When to start fresh:** long conversations cost more, because each message re-sends everything before it, and models get less focused as conversations grow. When you switch topics, use `/clear`. To keep going on the same task with less baggage, use `/compact`. It replaces the conversation with a summary the model writes, so later messages cost less. The tool suggests it when a conversation gets long.

## 8. Adjusting how it responds

| Setting | Type | Notes |
| --- | --- | --- |
| Instructions for every answer | `/system Always answer in plain language for a non-technical reader.` | `/system` alone shows the current instructions; `/system reset` restores the default. |
| Creativity | `/temperature 0.2` | Lower is more focused and consistent, higher is more varied. Range is 0 to 1. `/temperature default` resets it. |
| Answer length limit | `/max-tokens 1000` | Roughly 750 words. `/max-tokens default` removes the limit. |
| Streaming | `/stream off` | Shows each answer all at once instead of as it's written. |

Settings are remembered next time. Some models don't accept every setting; if so, the tool tells you and sends the message without that setting.

## 9. Using it in scripts

`-p` asks one question, prints only the answer and exits, so you can use it in other scripts or save the output:

```powershell
usai -p "write a one-paragraph summary of what a lease amendment is"
usai -p "summarize @weekly-report.txt" > summary.txt
Get-Content .\error.log | usai -p "what is the most likely cause of these errors?"
```

In this mode the model doesn't read files on its own (there's nobody to approve it), but `@` attachments and piped input work.

## 10. Keeping your API key current

USAi API keys expire every **7 days**. When yours expires, you'll see a `401` or `403` error. To fix it:

1. Rotate or create a key in the USAi Console (**Keys > My API keys**).
2. In USAi CLI, type `/key reset` and paste the new key.

Usage counts against a **weekly budget** that resets each Monday. If you hit it, you'll see a `429` error mentioning the budget.

## 11. Tips for better answers

- **Give context.** "Summarize this for my director, who has two minutes" gets a better result than "summarize this."
- **Show what you want.** Attach an example of a format you like and ask for the same style.
- **Attach the real thing.** Attach a file with `@` rather than pasting part of it; the model sees the whole thing.
- **Ask it to check itself.** "Double-check those numbers against the file" catches mistakes.
- **Verify before you rely on it.** AI models can be confidently wrong. Check facts, figures and anything about policy, contracts or law against an authoritative source.

## 12. Troubleshooting

**`usai` isn't recognized.** Open a *new* PowerShell window after installing. If it still isn't found, your PATH may be locked down; see "If you can't change your PATH" in the README.

**"Running scripts is disabled" or "cannot be loaded."** Unblock the file (see [Set it up](#2-set-it-up)). If that doesn't help, it's an organization policy; contact IT.

**`401` / `403`: key rejected.** Your key probably expired. See [Keeping your API key current](#10-keeping-your-api-key-current).

**`429`: too many requests or budget reached.** Rate limits clear in seconds, and the tool retries on its own. A budget limit resets on Monday.

**`503`: model unavailable.** The tool retries twice. If it still fails, wait and use `/retry`, or switch models with `/model`.

**The answer stopped partway with a note about the max-token limit.** `usai` already continued it twice for you. Say "continue," or raise the limit with `/max-tokens`.

**"Keep going?" after the model has taken 50 tool steps.** This is a check-in on long jobs. Press **y** to let it continue the same task, or **n** to stop it there.

**"File not found" when using `@`.** The name is looked up from your current folder. Check with `/pwd` and `/ls`, and put quotes around names with spaces.

**Strange characters in answers.** Use Windows Terminal (in the Microsoft Store, and preinstalled on newer Windows).

**Antivirus blocked the script.** Tell your IT or security team. Don't add an exclusion yourself.

## 13. Quick reference

```
START             usai              usai -c (continue)        usai -m <model>
ONE-OFF           usai -p "question"

ATTACH            @file.txt   @"name with spaces.pdf"   @folder\*.csv
MULTI-LINE        """ ... """   or paste   or end a line with \
STOP AN ANSWER    Esc
PREVIOUS INPUT    Up arrow

MODELS            /model   /model <part of name>   /models
CONVERSATIONS     /clear   /resume   /retry   /undo   /copy   /export   /context   /compact
FILES             /file <paths>   /drop   /cd <folder>   /ls   /pwd
TOOLS             /tools on|read|off   /approve ask|reads|edits|all   /revert   /permissions
                  y yes   a always this session   f always in this folder   n no (then say what instead)
                  push/deploy: type the word, every time
APPS SCRIPT       /clasp   /clasp <part of name>
INSTRUCTIONS      USAI.md in a folder (or ~\.usai)   /init to write one
SETTINGS          /system   /temperature   /max-tokens   /stream on|off
KEY               /key reset
UPDATE            usai -Update   or   /update
HELP / QUIT       /help   /exit
```
