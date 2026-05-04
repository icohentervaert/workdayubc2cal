# WorkdayUBC2Cal

A fast, privacy-first web tool that helps UBC students export their Workday course schedules into a standard calendar format.

## Why I Built This

After registering for courses on Workday, I wanted to see my classes in my daily calendar at a glance to make planning easier. Since Workday doesn't offer a built-in way to export courses to a calendar file, I decided to build a solution myself.

## Core Principles

- **Privacy First:** Your data is yours. All conversion logic happens entirely locally in your browser where no schedule or personal data is ever sent to a server.
- **Fast & Reliable:** Designed to load instantly and work seamlessly without unnecessary overhead.
- **Modern Development:** Built as an opportunity to experiment with new frameworks and evaluate AI-assisted development workflows.

## Getting Started (local)

To run this project locally, clone the repository and make sure you have [pnpm](https://pnpm.io/) installed. 

| Command                   | Action                                           |
| :------------------------ | :----------------------------------------------- |
| `pnpm install`            | Installs dependencies                            |
| `pnpm dev`                | Starts local dev server at `localhost:4321`      |
| `pnpm build`              | Build your production site to `./dist/`          |
| `pnpm preview`            | Preview your build locally, before deploying     |
| `pnpm astro ...`          | Run CLI commands like `astro add`, `astro check` |
| `pnpm astro -- --help`    | Get help using the Astro CLI                     |

## Development History

This project initially started out as a simple Python script. Wanting to make it more accessible to other students and experiment with web technologies, I decided to turn it into a web app.

I chose the **Astro** framework based on its great reputation and my prior experience with it, pairing it with **Tailwind CSS** to make styling easier.

For the core feature, converting the exported Workday Excel sheet, I initially tried using [PyScript](https://pyscript.net/) so I wouldn't have to rewrite my Python logic. However, after retrofitting it to work, I found the library's load times were just too slow for a good user experience.

To achieve the performance I wanted, I utilized an LLM to translate my Python script into pure JavaScript. From there, I used GitHub Copilot to help refine the UI into a clean, easy-to-use interface.

## What I Learned

Building this project was a great sandbox for:

- Evaluating the pros and cons of different web frameworks (like Astro vs. PyScript).
- Effectively leveraging LLMs to save time on specific, tedious tasks (like translating Python to JS).
- Using AI assistants for rapid frontend iteration and UI refinement.

## License

This project is licensed under the [MIT License](LICENSE) - see the LICENSE file for details.
