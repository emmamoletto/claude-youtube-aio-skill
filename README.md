# claude-youtube-ai-overviews-skill

Find out whether YouTube shows up in Google's AI Overviews for your sector, and whose videos get cited. A standalone [Claude skill](https://docs.claude.com/en/docs/agents-and-tools/agent-skills) that checks the Google SERP for a list of keywords, pulls the YouTube videos cited in the AI answer, looks up their channels with the YouTube Data API v3, and hands you an Excel report you can grow day by day.

> [!IMPORTANT]
> **This skill needs an SEO connector that returns Google SERPs with the AI Overview's source links.** It's built for the **SE Ranking MCP** and works with it out of the box. Using DataForSEO, Sistrix, Ahrefs, Semrush or another connector instead? See [Using a different SEO connector](#using-a-different-seo-connector). Without any such connector the skill can't run.

**Why:** AI Overviews cite YouTube more than almost any other domain, but that average hides big differences between sectors. Before you invest in video for AI search, you want to know three things for *your* queries: does the AI answer cite YouTube at all, which channels does it cite, and is that field open (lots of small channels) or crowded (two competitors with a hundred videos each)?

**Requirements:**

- A Claude that runs locally on your machine: [Claude Code](https://claude.com/claude-code) (CLI, desktop app or IDE extension) or the Claude desktop app. It won't work on claude.ai in the browser, because skills there run in a cloud sandbox that doesn't have your API key.
- An SEO connector (MCP) connected to Claude: SE Ranking by default, see above.
- A free YouTube Data API v3 key.
- Python 3 with `openpyxl`.

## What you get

One `.xlsx` file per audit, with seven tabs:

| Tab | What's in it |
|---|---|
| **Summary** | Totals for the whole period: keywords checked, % with an AI answer, % of those citing YouTube, appearances, unique videos, your channels' share, timestamped citations, share by channel type |
| **By day** | The same figures for each day collected |
| **Keywords** | One row per keyword: days checked, days with AI answer, days with YouTube, videos and channels cited, whether your channel was cited |
| **Detail** | One row per keyword per day: cited videos, channels, timestamps cited, organic YouTube results in the same SERP |
| **Videos** | One row per cited video: title, link, channel, channel type, duration, Long/Short, publish date, views, keywords, appearances, timestamped citations |
| **Channels** | One row per cited channel: handle, subscribers, appearances, keywords, **proposed type** (Creator, Publisher, Brand, Other) with a reason, and a **Confirmed type** dropdown for you |
| **Info** | List size, days collected, your channels, and a note on what "AI answer" covers |

Claude also writes a short summary in chat. Here's a real one, for 10 Italian personal-finance keywords:

> Day 1, 10 keywords. 9 have an AI answer (90%), and all 9 cite at least one YouTube video, for 22 appearances across 18 channels. No channel appears more than twice; the most cited are two lenders, a comparison site and an investing blog. 8 citations point to a timestamp, a sign that chapters matter. My reading, to confirm with more days: open field, because nobody dominates and the cited brand channels are small.

**Run it daily for about a week.** AI answers change from one day to the next, so one day is a weak sample. Each run adds a day to the same folder, and every report aggregates all the days collected so far.

Column headers and the summary can be in **English or Italian**: Claude picks the language you write in.

## Setup

### 1. Get a YouTube Data API v3 key (one time, ~5 minutes)

1. Create a project at [console.cloud.google.com](https://console.cloud.google.com) (free, no billing needed)
2. **APIs & Services → Library** → search **YouTube Data API v3** → **Enable**
3. **APIs & Services → Credentials → Create credentials → API key**
4. Click the new key → under **API restrictions** pick **Restrict key** → select **YouTube Data API v3** → **Save**. Treat the key like a password.

The free quota is 10,000 units a day. A run uses a handful (about one unit per 50 new videos), so quota is never the bottleneck.

### 2. Connect your SEO connector

**SE Ranking:** add the SE Ranking MCP to Claude following SE Ranking's MCP setup instructions (in your SE Ranking account, under API), and make sure your plan includes API units. Each keyword costs one advanced SERP (10 credits), so a 30-keyword list costs 300 credits a day. Claude tells you the cost and asks before spending anything.

**Another provider:** connect its MCP server the same way, then read [Using a different SEO connector](#using-a-different-seo-connector).

### 3. Download this repo

Click the green **Code** button at the top of this page → **Download ZIP**, and unzip it (or `git clone` if you're comfortable with git).

### 4. Install the skill, the easy way

Open Claude Code (or the Claude desktop app) in the unzipped folder and say:

> Install this skill into my Claude skills folder as "aio-youtube-field" and install its Python dependency. Then tell me how to set my YouTube API key as an environment variable, without writing the key into any file yourself.

Then set the key yourself (step 5), so it never passes through the chat.

<details>
<summary>Manual install (if you'd rather run the commands yourself)</summary>

From inside the unzipped folder:

```bash
# Copy the skill into Claude's skills folder
mkdir -p ~/.claude/skills/aio-youtube-field
cp -r SKILL.md scripts ~/.claude/skills/aio-youtube-field/

# Python dependency for the Excel file (try pip3 first; plain pip if that fails)
pip3 install openpyxl
```

</details>

### 5. Set your API key

The skill reads the key from the `YOUTUBE_API_KEY` environment variable. On macOS (zsh), add it to your shell profile:

```bash
echo 'export YOUTUBE_API_KEY="paste-your-key-here"' >> ~/.zshrc
```

Then **open a new terminal and restart Claude Code** so it picks up the variable. (On Linux with bash, use `~/.bashrc` instead.)

<details>
<summary>Alternative: set it only for one project</summary>

In the project folder where you run Claude Code, create `.claude/settings.local.json`:

```json
{
  "env": {
    "YOUTUBE_API_KEY": "paste-your-key-here"
  }
}
```

`settings.local.json` is meant to stay out of version control. Double-check it's in your `.gitignore` before committing anything.

</details>

### 6. Run it

In Claude Code, type:

> /aio-youtube-field

or just ask *"check whether YouTube shows up in AI Overviews for these keywords in the US market: …"*. Claude asks for anything missing:

- **Keywords**: 20 to 50 non-branded queries you want to appear for. No list? Ask Claude to build one for your sector.
- **Market**: language and country, e.g. `en` / `US` or `it` / `IT`.
- **Your channels** (optional): URLs, @handles or `UC...` IDs. Without them the report can't show your share.

Claude creates an audit folder in your working directory with the keyword list, one file per day, the raw SERPs and the `.xlsx`. **Tomorrow, run the same command and point it at the same folder** to add a day.

## Using a different SEO connector

The skill only needs one thing from the SEO tool: **a live Google SERP for a keyword in a given country and language, including the AI Overview with the URLs it cites.** Any connector that can return that will work.

- The SERP-ingest step is provider-agnostic. SE Ranking responses are parsed natively; anything else goes through a generic parser that finds the AI Overview block (`ai_overview` or `ai_mode`) and the YouTube links inside it, and Claude checks the first keyword by hand against the raw response.
- **DataForSEO**'s Google Organic SERP (advanced) returns the AI Overview with its references, so it's the closest drop-in. Depending on the endpoint you may have to ask it to load the AI Overview explicitly.
- **Ahrefs, Sistrix, Semrush and others**: check whether their MCP returns the AI Overview's *source links* for a live query. If it only reports that a keyword triggers an AI Overview, or only tracks whether *your* domain is cited, it isn't enough for this skill.

Paste this prompt into Claude Code to switch:

> I don't have the SE Ranking MCP. Use the aio-youtube-field skill with my **[DataForSEO / Sistrix / Ahrefs / …]** connector instead. First, look at that connector's tools and tell me which one returns a live Google SERP for a keyword, country and language **including the AI Overview and the URLs it cites**, what parameters it needs to load the AI Overview, and what one call costs. If none of its tools return the AI Overview's sources, stop and tell me. Otherwise, run the first keyword only, ingest the saved response with `--keyword`, then open the raw response and check by hand that the script found the AI Overview and every YouTube source in it. Show me the comparison and wait for my go-ahead before running the rest of the list. Keep every other step of the skill as it is.

If the check shows the parser misses something, Claude can adapt `scripts/aio_youtube_field.py` to the provider's format. The `parse_generic` function is the only part that would need to change.

Stick to one connector for the whole audit: providers fetch SERPs differently, so mixing them across days makes the days hard to compare.

## How it works

`scripts/aio_youtube_field.py` does the parsing, the YouTube lookups and the report. Claude runs the SERP calls and classifies the channels.

1. **init**: creates the audit folder with the keyword list, the market and your channels.
2. **status**: lists the keywords not yet checked today and the SERP calls (and SE Ranking credits) needed.
3. **ingest**: parses each saved SERP response, extracts the YouTube videos cited in the AI Overview (as a source or a link in the text) and any timestamp they're cited at, and stores one file per day. Raw responses are kept in `raw/`.
4. **resolve**: looks up every cited video and its channel with the YouTube API, cached so later days only look up new videos.
5. **channels**: prints the cited channels so Claude can propose a type for each (Creator, Publisher, Brand, Other) in `channel-types.json`.
6. **build**: writes the Excel file, aggregating every day collected.

The script also runs on its own (`python3 scripts/aio_youtube_field.py --help`).

## What it does and doesn't do

- It reads **Google's AI answer only**, the AI Overview block of a desktop SERP. It doesn't check ChatGPT, Perplexity, Gemini or other AI engines.
- The AI Overview block sometimes carries "AI Mode" in its text, so the report doesn't claim to separate AI Overviews from AI Mode.
- Channel types are Claude's **proposals**, based on the channel's name, description, size and cited videos. Confirm them in the Channels tab before you share the numbers.
- "Short" means the AI Overview cited a `/shorts/` URL; "Short (estimate)" means a video of 60 seconds or less cited with a regular URL.
- Every number comes from an ingested SERP or the YouTube API. Nothing is estimated or filled in.

## Security notes

- Your YouTube API key only grants read access to **public** YouTube data, but it's tied to your Google Cloud quota. Keep it out of repos and chats, and restrict it to the YouTube Data API v3 (step 1.4).
- The skill tells Claude never to write, print or log the key.
- SEO connector credentials stay in your MCP configuration; the skill never handles them.

## License

MIT
