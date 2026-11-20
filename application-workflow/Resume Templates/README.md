# Resume templates: pick the look you like

Module 03 can build your resume in seven layouts. The words are the same in every template; only the design changes. Each sample here is the same made-up candidate (Jordan Ellis, a data engineer at fictional companies) so you can compare the layouts side by side.

**Start with `00 - All templates at a glance.pdf`**: the first page of every template on two sheets. Open the numbered PDFs to see all three pages of one.

| # | Template | Name to put in your profile | Looks like | ATS |
|---|---|---|---|---|
| 01 | Classic | `classic` | Arial, centered name, bold underlined headings, one-line employer headers. The traditional default | Safe |
| 02 | Modern | `modern` | Calibri, name on the left in navy, thin rules under headings. Clean technology-leader look | Safe |
| 03 | Compact | `compact` | Arial with tight spacing and narrow margins. Holds the most text on three pages, for long careers | Safe |
| 04 | Executive | `executive` | Georgia serif, centered spaced-out headings between rules, dates on the right. Built for senior and executive roles | Safe |
| 05 | Two-Line | `two_line` | Cambria serif, company with dates on the right and the job title on its own line | Safe |
| 06 | Banner | `banner` | Name in a dark band across the top, shaded section bars, shaded Tech Stack strip | Safe |
| 07 | Hanging Headings | `hanging` | Section titles in a narrow left column beside their content; a consulting-style layout | Caution |

## How to choose

- **Applying through job portals and ATS systems:** any template marked Safe. They are single-column text with no tables, so every applicant tracking system reads them in the right order.
- **Senior, lead or manager roles:** Executive or Two-Line read as the most senior; Modern suits technology leaders.
- **A long career with lots to fit:** Compact.
- **Sending your resume straight to a person** (email, referral): any template, including Hanging Headings. It uses a table for its layout, which some older ATS read out of order, so avoid it for portal applications.

## How to use one

Open your Candidate Profile and set the line under RESUME LAYOUT, for example:

```
Resume template: executive
```

Leave it blank for Classic. Every new resume Module 03 builds uses that template and is fitted onto exactly three pages. To list the templates from a terminal: `python Tools/resume_render.py templates` (`python3` on a Mac).

You can change your mind at any time. New resumes use the new template; a resume reused from your library is rebuilt in your current template when it was made with templates (older resumes keep their original look).

## For maintainers

The samples are made from `sample_content.json` (all details are invented: fictional companies, a 555 phone number, an example.com email). After adding or changing a template in `Tools/resume_templates.json`, rebuild every PDF here with `python Tools/resume_samples.py`. It needs Microsoft Word (Windows) or LibreOffice, and PyMuPDF (`pip install pymupdf`) for the overview sheet.
