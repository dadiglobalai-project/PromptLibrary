# Dadi Coach Prompt Library Webpage

A Vercel-ready React/Vite webpage for the Dadi Coach Corporation ChatGPT Prompt Template Library.

## Included Features

- 1,000 official role-based system prompt templates from the Dadi Prompt Library document
- Search by task, title, category, role, output, or keyword
- Category, structure, department, and favorites filters
- Copy Prompt and Copy Full Card buttons
- Prompt Assistant for prompt recommendations
- Prompt Improvement Studio for rewriting rough prompts into role-based templates
- Upload custom prompts using JSON or CSV
- Save uploaded prompts in browser local storage
- Download all prompts as an Excel-compatible `.xls` file
- Download JSON backup
- Dadi Coach logo and green/yellow/clean theme options

## Required Repository Structure

```text
index.html
package.json
package-lock.json
vite.config.js
vercel.json
README.md
UPLOAD_INSTRUCTIONS.txt
public/
  dadi-coach-logo.png
  prompts.json
src/
  main.jsx
  App.jsx
```

## Vercel Settings

- Framework Preset: Vite
- Build Command: `npm run build`
- Output Directory: `dist`
- Install Command: default / leave blank

## Uploading Custom Prompts

The webpage supports uploading `.json` and `.csv` prompt files.

Recommended columns:

```text
ID, Title, Category, Structure, Expected Output, Department, Level, Best Use Case, Placeholders, Prompt
```

Uploaded prompts are stored only in the user's browser local storage. They are not saved to a server.

## Internal Use Reminder

Employees should verify AI-generated outputs before sending, publishing, or using them in official company materials.
