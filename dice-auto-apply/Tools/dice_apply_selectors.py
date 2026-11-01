"""Every Dice locator and every block-page text in one place (design section 8, "Selectors"). When Dice changes, only this file changes.

Each entry says when it was last checked against the real site:
  # verified 2026-10-30   seen on the live wizard with JOB-2539 (design section 2)
  # UNVERIFIED            written from the design; confirm in the first dry runs (plan, Phase 3) and change the comment to the date
Role- and text-based locators first; CSS only as a fallback. Playwright pierces open shadow DOM, which Dice's web components use.
"""
import re

# ---------- URLs ----------
JOB_URL = re.compile(r"^https://www\.dice\.com/job-detail/([0-9a-f-]{20,})", re.I)      # verified 2026-10-30
# verified 2026-11-01: the wizard now opens IN PLACE on /job-detail/<guid> (the URL does not change). The flow recognises the wizard by
# STEP_RESUME_TEXT and success by SUCCESS_TEXT; these URL patterns are kept for the older behaviour seen on 2026-10-30.
WIZARD_URL = re.compile(r"/job-applications/[0-9a-f-]+/wizard", re.I)                  # verified 2026-10-30 (older behaviour)
SUCCESS_URL = re.compile(r"/wizard/success", re.I)                                      # verified 2026-10-30 (older behaviour)
HOME_URL = "https://www.dice.com"

# ---------- texts ----------
STEP_RESUME_TEXT = "Resume & Cover Letter"                                              # verified 2026-11-01
STEP_REVIEW_TEXT = "Review your application"                                            # verified 2026-10-30
SUCCESS_TEXT = "on its way"                                                             # verified 2026-10-30 ("Your application is on its way!")
NEW_FILE_TEXT = "New file"                                                              # verified 2026-10-30 (shown on the card after Replace)
GONE_TEXTS = ["no longer available", "no longer accepting applications", "this job has expired", "job has expired",
              "position has been filled"]                                               # UNVERIFIED

# Block pages (hard signals). Generic phrases are matched only in the page title and headings, because a job description can say
# "monitor unusual activity" or "access denied errors" (security and IAM roles). The body of a short page is checked only for phrases
# that appear on bot-check pages and nowhere else.
BLOCK_TEXTS = ["just a moment", "access denied", "unusual activity", "verify you are human", "are you a robot",
               "press & hold", "press and hold", "request blocked", "too many requests", "checking your browser"]  # UNVERIFIED
BLOCK_BODY_TEXTS = ["verify you are human", "are you a robot", "press & hold", "press and hold", "checking your browser",
                    "request blocked", "enable javascript and cookies to continue"]                                  # UNVERIFIED
SHORT_PAGE_CHARS = 2500
CAPTCHA_FRAME_PATTERNS = [r"recaptcha", r"hcaptcha", r"challenges\.cloudflare\.com", r"turnstile", r"captcha-delivery",
                          r"perimeterx|px-captcha", r"arkoselabs|funcaptcha"]           # UNVERIFIED
# Matched in titles/headings, or in the body of a short page only (never in a long job description). UNVERIFIED (open question 6)
DICE_LIMIT_TEXTS = ["reached the maximum number of applications", "daily application limit", "application limit reached",
                    "you have reached your limit"]


# ---------- locators (each takes the Playwright page) ----------
APPLY_NAME = re.compile(r"^\s*((easy\s+)?apply(\s+now)?|continue\s+application)\s*$", re.I)


def apply_button(page):
    # verified 2026-11-01: "Apply" (button) when logged in. After a wizard was left unfinished, Dice saves a draft and shows
    # "Continue Application" (a link to /job-applications/<guid>/wizard, opening at Step 1 with the draft's resume) — verified 2026-11-01.
    # "Apply Now" / "Easy Apply" also reported (open question 3).
    return page.get_by_role("button", name=APPLY_NAME).or_(page.get_by_role("link", name=APPLY_NAME)).first


def applied_badge(page):
    return page.get_by_text(re.compile(r"^\s*applied\s*$", re.I)).first                 # UNVERIFIED


def login_link(page):
    return page.get_by_role("link", name=re.compile(r"^\s*(log\s*in|sign\s*in)\s*$", re.I)).first   # UNVERIFIED


def logged_in_marker(page):
    # a profile/avatar menu in the header                                                  UNVERIFIED
    return page.locator("[data-cy='header-profile'], [aria-label*='profile' i], [aria-label*='account' i], img[alt*='avatar' i]").first


def file_options_button(page):
    return page.get_by_role("button", name=re.compile(r"file options", re.I)).first     # verified 2026-11-01 (aria-label "File options")


def replace_menuitem(page):
    return page.get_by_role("menuitem", name=re.compile(r"^\s*replace\s*$", re.I)).first  # verified 2026-10-30


def resume_upload_zone(page):
    # the resume upload area when the profile has no resume                                UNVERIFIED
    return page.get_by_role("button", name=re.compile(r"upload.*resume|resume.*upload", re.I)).first


def next_button(page):
    return page.get_by_role("button", name=re.compile(r"^\s*next\s*$", re.I)).first     # verified 2026-10-30


def submit_button(page):
    return page.get_by_role("button", name=re.compile(r"^\s*submit\s*$", re.I)).first   # verified 2026-10-30


def wizard_form(page):
    return page.locator("form, main").first                                              # UNVERIFIED


# Never used by the flow, listed so a reviewer can see they are deliberately avoided (design section 2 and 8).
NEVER_CLICK = ["Delete", "Back", "Cancel"]
