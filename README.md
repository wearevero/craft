# Craft

Craft is a monorepo of automation tools designed to handle repetitive daily tasks efficiently. Each tool lives in its own directory and is built with simplicity and configurability in mind, focused on automating data extraction and processing workflows so they run consistently without manual intervention.

## Purpose

Manual, repetitive work doesn't scale and it doesn't stay accurate forever — copy-pasting the same report every day is exactly the kind of task that eventually gets a step skipped or a number transposed. Craft exists to take that class of work off your plate: scripted, repeatable, and predictable, so the people who used to do it by hand can spend that time on things that actually need judgment.

As the number of tools grows, this repo is meant to be the single place they all live — rather than scattered across individual repos, each reinventing its own conventions.

## Project Structure

Craft is organized as a collection of independent tools, each living in its own top-level directory. A tool is typically responsible for a single workflow end-to-end — extracting data from a source, processing it, and producing an output artifact (a spreadsheet, a file, a report, etc.).

```
craft/
├── invoice/       # Invoice-related automation
├── take-file/     # File retrieval/processing tool
├── wip/           # Table-data scraper and Excel exporter
└── README.md      # You are here
```

New tools are added as new top-level directories following the same pattern. Each tool directory contains its own `README.md` with setup instructions, required environment variables, and usage details specific to that tool. This root README covers the monorepo as a whole; tool-level READMEs cover the specifics.
