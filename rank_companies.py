"""
Detect relevant emails, then rank them. Built for JOB APPLICATIONS by default.

Application stages detected from your emails (furthest stage reached wins):

    Applied -> Under Review -> Assessment -> Interview -> Pending After Interview
            -> Offer      (or Rejected at any point)

Two modes share one engine:
    MODE = "job"       job applications (default)
    MODE = "sponsor"   sponsorship outreach (the original use)

Switch by editing MODE below, or per run without editing:
    Windows cmd:   set RANKER_MODE=sponsor
    PowerShell:    $env:RANKER_MODE="sponsor"

    pip install vaderSentiment
    python rank_companies.py

Pipeline:
    emails -> relevant? (keyword score + thread propagation)
    -> per company: intent (regex) -> stage -> sentiment (VADER) -> email score
    -> company score (fit + history) -> base score -> intent multiplier
    -> recency decay -> final score -> rank within intent group
"""

import csv
import json
import os
import re
from collections import defaultdict
from datetime import datetime, timezone

from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

# ---------------------------------------------------------------------------
# CONFIG: edit these
# ---------------------------------------------------------------------------
MODE = os.environ.get("RANKER_MODE", "job")   # "job" or "sponsor"

MY_EMAIL = "you@gmail.com"              # your address, so your own mail isn't scored
CACHE_FILE = "emails_cache.json"
OUTPUT_CSV = "ranked_companies.csv"
IGNORE_DOMAINS = {"gmail.com", "googlemail.com", "google.com", "accounts.google.com"}
# Your own address's domain is ignored automatically, so an Outlook mailbox
# doesn't count outlook.com as a company. Add other personal domains here if needed.

EMAIL_WEIGHT, COMPANY_WEIGHT = 0.6, 0.4   # matches your example (76 = .6*80 + .4*70)
USE_RECENCY = True                        # lose 10% per 30 days of silence, floor 60%
FOLLOWUP_DAYS = 10                        # default quiet period before "follow up"
SPONSOR_THRESHOLD = 3                     # score needed to count as a relevant email

# Job mode: how long to wait at each stage before flagging a follow-up
FOLLOWUP_BY_STAGE = {"Applied": 10, "Responded": 7, "Under Review": 14,
                     "Assessment": 5, "Interview": 5, "Pending After Interview": 5}
# Job mode, optional: if an interview invite is this many days old with nothing newer,
# assume the interview happened and mark "Pending After Interview". None = only use wording.
PENDING_AFTER_DAYS = None

# intent -> (group order, multiplier). Lower order = higher priority group.
INTENTS = {
    "Strong Interest": (1, 1.00),
    "Interested":      (2, 0.80),
    "Considering":     (3, 0.60),
    "Unclear":         (4, 0.40),
    "Unknown":         (5, 0.20),
    "Not Interested":  (6, 0.00),
}

FIT_SCORE = {"High": 85, "Medium": 60, "Low": 35}   # starting guesses, tune later
HISTORY_BONUS = 15      # added if "history": True (referral, prior interview, sponsored before)

# Fill these in by hand. Unlisted companies get Medium fit, no history.
COMPANY_META = {
    # "examplecorp.com": {"name": "Example Corp", "fit": "High", "history": True},
    # "acme corp":       {"fit": "High"},        # portal emails are keyed by sender name
}

# Manual corrections always win over the regex rules.
# stage:  job mode -> Applied / Under Review / Assessment / Interview /
#                     Pending After Interview / Offer / Rejected
# intent: Strong Interest / Interested / Considering / Unclear / Unknown / Not Interested
#         (or the display names shown in the dashboard)
OVERRIDES = {
    # "examplecorp.com": {"stage": "Pending After Interview"},
}

# ---------------------------------------------------------------------------
# JOB KEYWORDS (edit these to match the emails you actually receive)
# ---------------------------------------------------------------------------
I = re.I

# Offer
JOB_OFFER = re.compile(
    r"pleased to offer|(we|i)(?:'|’)d (like|love) to (offer|extend)|we would like to (offer|extend)|"
    r"offer of employment|extend (you )?an offer|offer letter|formal offer|"
    r"congratulations[^.]{0,80}(offer|join|position|role)", I)

# Rejection
JOB_REJECT = re.compile(
    r"regret to inform|sorry to inform|we regret|"
    r"unfortunately[^.]{0,120}(other candidates|not (be )?(moving|proceeding|going) forward|"
    r"not (been )?selected|position (has been|was) filled|unable to (offer|move forward|proceed|extend)|"
    r"decided (not )?to)|"
    r"not (be )?(moving|proceeding|going) forward|will not be (moving|proceeding)|"
    r"won(?:'|’)?t be (moving|proceeding)|"
    r"(decided|chosen|chose|elected) (not )?to (move|proceed|go|pursue)( forward)?( with)? (other|another)|"
    r"(pursue|moving forward with|proceed with) (other|another) candidates?|"
    r"not (been )?selected|position (has been|was) filled|"
    r"no longer (considering|being considered|accepting)|"
    r"after careful (consideration|review)[^.]{0,80}\b(not|unable|other)\b|"
    r"not (a|the) (good |right )?fit|unable to (offer|move forward|proceed|extend)|"
    r"(did|do) not (move|proceed) forward", I)

# Interview invited / scheduled
JOB_INTERVIEW = re.compile(
    r"invite you to (an? |the |our )?(interview|next round|onsite|on-site|phone screen|screen|final round|panel)|"
    r"schedule (an? |your |the )?(interview|phone screen|call|chat|conversation)|"
    r"interview (on|at|scheduled|invitation|confirmation|details|slot)|your interview|"
    r"(onsite|on-site|final|panel|technical|virtual) interview|phone screen|"
    r"calendar invite|zoom\.us|meet\.google\.com|teams\.microsoft\.com|calendly\.com|"
    r"select a (time|slot)|book a time|pick a time|"
    r"would like to (speak|talk|chat|meet) with you|"
    r"we(?:'|’)d (love|like) to (speak|talk|chat|meet|interview)", I)

# After the interview, waiting on a decision (needs an earlier interview message)
JOB_POST_INTERVIEW = re.compile(
    r"thank you for (interviewing|the interview|taking the time to (meet|speak|interview|talk|chat)|"
    r"meeting with|speaking with|your time today)|"
    r"enjoyed (speaking|meeting|talking|our (conversation|chat))|"
    r"(great|nice|a pleasure) (speaking|meeting|talking|chatting)( with you)?|"
    r"following up (on|after) (our|my|the) (interview|conversation|chat|call)|"
    r"after (your|the|our) (interview|conversation|call|chat)|"
    r"(still|currently) (interviewing|speaking with|meeting with) (other )?candidates|"
    r"finaliz(e|ing) (our|the) (decision|candidates|interviews|hiring)|"
    r"(decision|update)[^.]{0,40}(by|next week|in the coming|within|soon|shortly)|"
    r"(will|we(?:'|’)ll) (be in touch|get back to you|reach out|let you know)[^.]{0,60}"
    r"(decision|after|once|next week|in the coming|following)|"
    r"in the coming (days|weeks)|next steps (will|to) follow", I)

# Assessment / take-home
JOB_ASSESS = re.compile(
    r"(complete|take|fill out|submit|finish)[^.]{0,60}"
    r"(assessment|questionnaire|coding challenge|take-home|hackerrank|codesignal|online test)|"
    r"(assessment|coding challenge|take-home)[^.]{0,40}(link|attached|below|instructions)", I)

# Application is being reviewed
JOB_REVIEW = re.compile(
    r"under review|currently reviewing|(is|are|being|been) reviewed|"
    r"reviewing (your |all )?(application|applications|candidates|resumes?)|"
    r"(will|we(?:'|’)ll|team will|committee will) (review|be reviewing)|"
    r"review process|being considered", I)

# Recruiter reaching out / positive signal (no specific stage yet)
JOB_POSITIVE = re.compile(
    r"send (us|me) (your )?(availability|references|portfolio|transcript|resume)|"
    r"(like|love|want) to learn more about you|impressed (by|with)|strong (background|profile)|"
    r"(background|profile|experience) (looks|seems|is) (great|strong|impressive|interesting)|"
    r"interested in (your|speaking|talking|connecting)|your availability", I)


def _either(*patterns):
    return re.compile("|".join(p.pattern for p in patterns), I)


# ---------------------------------------------------------------------------
# MODE PROFILES: everything that differs between job hunting and sponsorship
# ---------------------------------------------------------------------------
APOS = "['\u2019]?"

PROFILES = {
    "job": {
        "name": "job application",
        "match_label": "Job-related", "no_label": "Not job-related",
        "titles": ("PRIORITY LIST (active applications)", "OFFERS", "REJECTED"),
        "query": ('newer_than:180d (application OR "thank you for applying" OR interview '
                  'OR recruiter OR "next steps" OR "your candidacy" OR "offer letter" '
                  'OR assessment OR "hiring team")'),
        "core_label": "application wording",
        "core": re.compile(
            r"thank you for (applying|your application|your interest)|(your|my) application|"
            r"application (for|to|has been|was|status|received)|applied (for|to)|"
            r"\binterview(s|ing)?\b|\brecruiter\b|hiring (team|manager)|"
            r"job (offer|opening|application)|offer (letter|of employment)|candidacy|"
            r"next steps|phone screen|coding challenge|take-home|"
            r"(opportunity|role|position|opening) (at|with|for)\b|"
            r"your (background|profile|experience) (looks|seems)", I),
        "support": re.compile(
            r"\b(resume|cv|position|role|opening|salary|compensation|start date|"
            r"background check|references|availability|onsite|on-site|assessment|"
            r"coding challenge|take-home|skills|qualifications|candidate|interviewer|"
            r"hiring manager|final round|panel|offer)\b", I),
        # ATS confirmations come from noreply addresses, so only penalize job-alert mail
        "junk_body": re.compile(
            r"job alerts?|jobs? (you may|recommended for)|recommended jobs|similar jobs|"
            r"new jobs? (for you|matching|posted)|"
            r"you are receiving this (email|message) because you (created|signed up|subscribed)", I),
        "junk_sender": re.compile(
            r"job-?alerts?|jobs-listings|alert@indeed|jobs@(linkedin|indeed|ziprecruiter|glassdoor)|"
            r"newsletter|digest@", I),
        "rules": [
            ("Not Interested", JOB_REJECT),
            ("Strong Interest", _either(JOB_OFFER, JOB_INTERVIEW)),
            ("Interested", _either(JOB_ASSESS, JOB_POSITIVE)),
            ("Considering", JOB_REVIEW),
        ],
        "meeting": JOB_INTERVIEW,
        "win_pattern": JOB_OFFER,
        "stages": {"not_contacted": "Not Applied", "contacted": "Applied",
                   "responded": "Responded", "review": "Under Review",
                   "assessment": "Assessment", "meeting": "Interview",
                   "pending": "Pending After Interview", "win": "Offer", "lose": "Rejected"},
        # interest tier implied by each stage (stages not listed use the keyword intent)
        "stage_tier": {"Offer": "Strong Interest", "Interview": "Strong Interest",
                       "Pending After Interview": "Interested", "Assessment": "Interested",
                       "Under Review": "Considering", "Rejected": "Not Interested"},
        "intent_display": {"Strong Interest": "Interview / Offer", "Interested": "Advancing",
                           "Considering": "Under Review", "Unclear": "Acknowledged",
                           "Unknown": "No Response", "Not Interested": "Rejected"},
        "reply_needed": ("Strong Interest", "Interested"),   # auto-acks don't need your reply
        "win_status": "Offer - respond",
        "pending_stage": "Pending After Interview",
        # Applicant-tracking systems send mail for many employers: key those by sender NAME
        "ats_domains": {"greenhouse.io", "greenhouse-mail.io", "lever.co", "myworkday.com",
                        "myworkdayjobs.com", "icims.com", "smartrecruiters.com", "ashbyhq.com",
                        "jobvite.com", "taleo.net", "workable.com", "bamboohr.com",
                        "successfactors.com", "breezy.hr", "recruitee.com", "applytojob.com",
                        "jazzhr.com"},
    },
    "sponsor": {
        "name": "sponsorship",
        "match_label": "Sponsor", "no_label": "Not sponsor",
        "titles": ("PRIORITY LIST (active)", "SPONSORED (manage)", "REJECTED"),
        "query": "in:sent (sponsor OR sponsorship OR partnership) newer_than:365d",
        "core_label": "sponsor word",
        "core": re.compile(r"\bsponsor(s|ship|ships|ed|ing)?\b", I),
        "support": re.compile(
            r"\b(partnership|partner with|in-?kind|donation|media kit|package|tier|"
            r"logo|booth|brand exposure|fundrais\w+|our (event|mission|programs?)|"
            r"support (us|our))\b", I),
        "junk_body": re.compile(
            r"unsubscribe|view (this )?(email )?in (your )?browser|manage (your )?preferences", I),
        "junk_sender": re.compile(r"no-?reply|do-?not-?reply|mailer-daemon|notifications?@", I),
        "rules": [
            ("Not Interested", re.compile(
                r"not interested|unable to (sponsor|support|participate)|"
                rf"can{APOS}t (sponsor|support)|cannot (sponsor|support)|"
                r"no longer (sponsor|accepting)|not (able|in a position) to|"
                r"not a (good )?fit|not at this time|decline|pass on this|"
                r"budget (is|has been) (full|allocated|exhausted)", I)),
            ("Strong Interest", re.compile(
                rf"we(?:{APOS}d| would) (love|be happy|be glad|like) to sponsor|"
                r"happy to (sponsor|move forward|proceed)|"
                r"would like to (move forward|sponsor)|let.?s move forward|count us in|"
                r"ready to (sponsor|proceed)|send (over )?(the )?(invoice|payment|contract|agreement)",
                I)),
            ("Interested", re.compile(
                r"(more|additional) (info|information|details)|sponsorship (packages?|levels|tiers|deck)|"
                r"media kit|tell me more|what (are|is) (the )?(sponsorship|benefits|cost|price)|"
                r"(send|share) (us |me )?(over )?(your )?(info|information|details|package)", I)),
            ("Considering", re.compile(
                r"discuss (this |it )?(internally|with (the|my) (team|board|manager))|"
                r"review(ing)? (this |it )?(with|internally)|get back to you|circle back|"
                r"bring (this|it) to|under review|evaluat(e|ing)|consider(ing)?|check with", I)),
        ],
        "meeting": re.compile(
            r"(schedule|set up|book).{0,30}(call|meeting|zoom|chat)|calendar invite|"
            r"are you (free|available)|zoom\.us|meet\.google\.com", I),
        "win_pattern": None,                    # "Sponsored" is set by hand in OVERRIDES
        "stages": {"not_contacted": "Not Contacted", "contacted": "Contacted",
                   "responded": "Responded", "discussion": "In Discussion",
                   "meeting": "Meeting Scheduled", "win": "Sponsored", "lose": "Declined"},
        "stage_tier": {},
        "intent_display": {},
        "reply_needed": ("Strong Interest", "Interested", "Considering", "Unclear"),
        "win_status": None,
        "pending_stage": None,
        "ats_domains": set(),
    },
}
PROFILE = PROFILES[MODE]
STAGES = PROFILE["stages"]
INTENT_DISPLAY = PROFILE["intent_display"]
INTENT_REVERSE = {v: k for k, v in INTENT_DISPLAY.items()}


# ---------------------------------------------------------------------------
# STEP 1: Is this email relevant? (job application / sponsorship)
# ---------------------------------------------------------------------------
def sponsor_score(e):
    """Return (score, reasons) for one email dict."""
    P = PROFILE
    subject, body = e.get("subject", ""), e.get("body", "")[:3000]
    score, reasons = 0, []

    if P["core"].search(subject):
        score += 4
        reasons.append(f"{P['core_label']} in subject")
    n_core = len(P["core"].findall(body))
    if n_core:
        score += min(3, n_core)
        reasons.append(f"{P['core_label']} in body x{n_core}")

    support = {m.group(0).lower() for m in P["support"].finditer(subject + " " + body)}
    if support:
        score += min(3, len(support))
        reasons.append("related terms: " + ", ".join(sorted(support)[:4]))

    if P["junk_sender"].search(e.get("sender_email", "")):
        score -= 4
        reasons.append("automated or alert sender")
    if P["junk_body"].search(body):
        score -= 3
        reasons.append("newsletter/alert footer")
    return score, reasons


def is_sponsor_email(e):
    score, reasons = sponsor_score(e)
    return score >= SPONSOR_THRESHOLD, score, reasons


def filter_sponsor_emails(emails):
    """Keep relevant emails, plus every message in the same thread.

    Short replies like "Sounds good, see you then" never use the keywords, so one
    matching message pulls its whole thread (or, without thread IDs, its
    sender domain) into the results.
    """
    hit_threads = {e["thread_id"] for e in emails
                   if e.get("thread_id") and is_sponsor_email(e)[0]}
    hit_domains = {e["domain"] for e in emails
                   if not e.get("thread_id") and e["domain"] not in IGNORE_DOMAINS
                   and is_sponsor_email(e)[0]}
    return [e for e in emails
            if e.get("thread_id") in hit_threads
            or (not e.get("thread_id") and e["domain"] in hit_domains)
            or is_sponsor_email(e)[0]]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
analyzer = SentimentIntensityAnalyzer()


def parse_date(s):
    dt = datetime.fromisoformat(s.replace("Z", "+00:00"))
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


def classify_intent(text):
    for name, pattern in PROFILE["rules"]:
        if pattern.search(text):
            return name
    return None


def message_text(e):
    return f"{e.get('subject', '')}\n{e.get('body', '')[:2000]}"


def sentiment_label(compound):
    if compound is None:
        return "Unknown"
    return "Positive" if compound >= 0.25 else "Negative" if compound <= -0.25 else "Neutral"


def company_key(e):
    """Domain, except applicant-tracking systems, which are keyed by sender name."""
    d = e["domain"]
    if e.get("sender_name") and any(d == a or d.endswith("." + a) for a in PROFILE["ats_domains"]):
        name = re.sub(r"\b(recruiting|recruitment|careers?|talent( acquisition)?|hiring( team)?|"
                      r"jobs?|hr|team|notifications?|no-?reply|via \w+)\b", "",
                      e["sender_name"], flags=re.I)
        name = re.sub(r"[^\w &.-]", "", name).strip(" .-").lower()
        if name:
            return name
    return d


def group_by_company(emails):
    """Group messages by company, using THREADS so replies from a different
    address or colleague still count toward the company you wrote to."""
    mine_addr = MY_EMAIL.lower()
    skip = IGNORE_DOMAINS | {mine_addr.split("@")[-1]}

    def from_me(e):
        return e["sender_email"].lower() == mine_addr or "SENT" in e.get("labels", [])

    def recipients(e):
        return [d for d in re.findall(r"@([\w.-]+)", e.get("to", "").lower()) if d not in skip]

    # thread -> company domain, taken from who WE wrote to first
    thread_company = {}
    for e in sorted(emails, key=lambda e: e.get("date") or ""):
        tid = e.get("thread_id")
        if tid and from_me(e) and tid not in thread_company and recipients(e):
            thread_company[tid] = recipients(e)[0]

    companies = defaultdict(list)
    for e in emails:
        mine = from_me(e)
        if e.get("thread_id") in thread_company:
            domains = {thread_company[e["thread_id"]]}
        elif mine:
            domains = set(recipients(e))
        else:
            domains = {company_key(e)} - skip
        for d in domains:
            companies[d].append((e, mine))
    return companies


# ---------------------------------------------------------------------------
# STEP 2: Work out the stage (job and sponsor modes differ)
# ---------------------------------------------------------------------------
def stage_sponsor(items, inbound, outbound, intent, days_since):
    S = STAGES
    if intent == "Not Interested":
        return S["lose"]
    if any(PROFILE["meeting"].search(message_text(e)) for e in inbound):
        return S["meeting"]
    if len(inbound) >= 2:
        return S["discussion"]
    if len(inbound) == 1:
        return S["responded"]
    return S["contacted"] if outbound else S["not_contacted"]


def stage_job(items, inbound, outbound, intent, days_since):
    """Furthest stage reached. Offer and Rejected beat everything else."""
    S = STAGES
    msgs = sorted(((parse_date(e["date"]), mine, message_text(e))
                   for e, mine in items if e["date"]), key=lambda m: m[0])
    theirs = [t for _, mine, t in msgs if not mine]

    if any(JOB_OFFER.search(t) for t in theirs):
        return S["win"]
    if intent == "Not Interested":
        return S["lose"]

    interview_dates = [d for d, _, t in msgs if JOB_INTERVIEW.search(t)]
    if interview_dates:
        first = min(interview_dates)
        # post-interview wording (from either side) in a LATER message than the invite
        if any(d > first and JOB_POST_INTERVIEW.search(t) for d, _, t in msgs):
            return S["pending"]
        if PENDING_AFTER_DAYS is not None and days_since is not None and days_since >= PENDING_AFTER_DAYS:
            return S["pending"]
        return S["meeting"]
    if any(JOB_ASSESS.search(t) for t in theirs):
        return S["assessment"]
    if any(JOB_REVIEW.search(t) for t in theirs):
        return S["review"]
    if inbound:
        return S["responded"] if intent in ("Strong Interest", "Interested") else S["contacted"]
    return S["contacted"] if outbound else S["not_contacted"]


STAGE_FN = {"job": stage_job, "sponsor": stage_sponsor}[MODE]


# ---------------------------------------------------------------------------
# STEP 3: Analyze one company
# ---------------------------------------------------------------------------
def analyze_company(domain, items, now):
    meta = COMPANY_META.get(domain, {})
    override = OVERRIDES.get(domain, {})

    inbound = sorted((e for e, mine in items if not mine), key=lambda e: e["date"] or "")
    outbound = sorted((e for e, mine in items if mine), key=lambda e: e["date"] or "")
    all_dates = [parse_date(e["date"]) for e, _ in items if e["date"]]
    last_contact = max(all_dates) if all_dates else None
    days_since = (now - last_contact).days if last_contact else None

    # --- keyword intent: newest inbound message with a recognizable signal wins
    intent = None
    for e in reversed(inbound):
        intent = classify_intent(message_text(e))
        if intent:
            break
    if intent is None:
        intent = "Unclear" if inbound else "Unknown"

    # --- stage, then the interest tier that stage implies
    stage = override["stage"] if "stage" in override else STAGE_FN(
        items, inbound, outbound, intent, days_since)
    intent = PROFILE["stage_tier"].get(stage, intent)
    if "intent" in override:
        intent = INTENT_REVERSE.get(override["intent"], override["intent"])

    # --- email score (0-100): average VADER sentiment of last 3 inbound messages
    compounds_all = [analyzer.polarity_scores(message_text(e))["compound"] for e in inbound]
    compounds = compounds_all[-3:]
    avg = sum(compounds) / len(compounds) if compounds else None
    email_score = (avg + 1) / 2 * 100 if avg is not None else 50.0   # neutral if no reply

    # --- company score (0-100): fit + history bonus
    fit = meta.get("fit", "Medium")
    history = bool(meta.get("history", meta.get("previous_sponsor")))
    company_score = min(100, FIT_SCORE[fit] + (HISTORY_BONUS if history else 0))

    # --- base, multiplier, recency, final
    base = EMAIL_WEIGHT * email_score + COMPANY_WEIGHT * company_score
    order, multiplier = INTENTS[intent]
    decay = 1.0
    if USE_RECENCY and days_since is not None:
        decay = max(0.6, 1 - 0.10 * (days_since / 30))
    final = base * multiplier * decay

    # --- conversation health (thread view)
    sent_label = sentiment_label(avg)
    thread_ids = {e.get("thread_id") or e.get("id") for e, _ in items}
    replied_threads = {e.get("thread_id") or e.get("id") for e in inbound}

    response_days = None
    if outbound and outbound[0]["date"]:
        first_out = parse_date(outbound[0]["date"])
        replies = [parse_date(e["date"]) for e in inbound
                   if e["date"] and parse_date(e["date"]) > first_out]
        if replies:
            response_days = round((min(replies) - first_out).total_seconds() / 86400, 1)

    dated = sorted((parse_date(e["date"]), mine) for e, mine in items if e["date"])
    last_from_them = bool(dated) and not dated[-1][1]

    trend = "n/a"
    if len(compounds_all) >= 2:
        delta = compounds_all[-1] - compounds_all[0]
        trend = "Warming" if delta >= 0.15 else "Cooling" if delta <= -0.15 else "Steady"

    stale = days_since is not None and days_since > FOLLOWUP_BY_STAGE.get(stage, FOLLOWUP_DAYS) \
        if MODE == "job" else (days_since is not None and days_since > FOLLOWUP_DAYS)
    if stage == STAGES["win"]:
        convo_status = PROFILE["win_status"] or stage
    elif stage == STAGES["lose"] or intent == "Not Interested":
        convo_status = STAGES["lose"]
    elif stage == PROFILE["pending_stage"]:
        convo_status = "Follow up on decision" if stale else "Waiting on decision"
    elif not inbound:
        convo_status = "No reply - follow up" if stale else "Awaiting first reply"
    elif last_from_them and intent in PROFILE["reply_needed"]:
        convo_status = "Reply needed from you"
    else:
        convo_status = "Stalled - follow up" if stale else "Waiting on them"

    if stage == STAGES["win"]:
        health = "Won"
    elif intent == "Not Interested":
        health = "Lost"
    elif not inbound:
        health = "No data"
    elif trend == "Cooling" or sent_label == "Negative":
        health = "Cooling"
    elif intent in ("Strong Interest", "Interested", "Considering"):
        health = "Going well"
    else:
        health = "Mixed"

    active = stage not in (STAGES["win"], STAGES["lose"]) and intent != "Not Interested"
    pretty = domain.title() if "." not in domain else domain
    return {
        "company": meta.get("name", pretty),
        "domain": domain,
        "intent": INTENT_DISPLAY.get(intent, intent),
        "stage": stage,
        "fit": fit,
        "sentiment": sent_label,
        "history": history,
        "last_contact": last_contact.date().isoformat() if last_contact else "",
        "days_since": days_since,
        "email_score": round(email_score, 1),
        "company_score": company_score,
        "base_score": round(base, 1),
        "multiplier": multiplier,
        "recency_decay": round(decay, 2),
        "final_score": round(final, 1),
        "threads": len(thread_ids),
        "threads_replied": len(replied_threads),
        "response_days": response_days,
        "last_speaker": "Them" if last_from_them else "You",
        "trend": trend,
        "health": health,
        "convo_status": convo_status,
        "follow_up": "follow up" in convo_status.lower(),
        "_order": order,
        "_list": "won" if stage == STAGES["win"] else "lost" if not active else "active",
    }


# ---------------------------------------------------------------------------
# STEP 4: Rank and report
# ---------------------------------------------------------------------------
def rank(emails):
    now = datetime.now(timezone.utc)
    results = [analyze_company(d, items, now) for d, items in group_by_company(emails).items()]
    active = sorted((r for r in results if r["_list"] == "active"),
                    key=lambda r: (r["_order"], -r["final_score"]))   # intent first, then score
    won = [r for r in results if r["_list"] == "won"]
    lost = [r for r in results if r["_list"] == "lost"]
    return active, won, lost


def print_table(title, rows):
    print(f"\n=== {title} ({len(rows)}) ===")
    if not rows:
        print("  (none)")
        return
    print(f"{'#':<3}{'Company':<22}{'Intent':<18}{'Stage':<25}{'Health':<12}"
          f"{'Convo status':<24}{'Final':>6}  Last")
    for i, r in enumerate(rows, 1):
        print(f"{i:<3}{r['company'][:21]:<22}{r['intent']:<18}{r['stage']:<25}{r['health']:<12}"
              f"{r['convo_status']:<24}{r['final_score']:>6}  {r['last_contact']}")


if __name__ == "__main__":
    with open(CACHE_FILE, encoding="utf-8") as f:
        all_emails = json.load(f)

    emails = filter_sponsor_emails(all_emails)
    print(f"Mode: {MODE} | {len(emails)} of {len(all_emails)} emails look "
          f"{PROFILE['name']}-related")

    active, won, lost = rank(emails)
    t_active, t_won, t_lost = PROFILE["titles"]
    print_table(t_active, active)
    print_table(t_won, won)
    print_table(t_lost, lost)

    rows = active + won + lost
    cols = [k for k in rows[0] if not k.startswith("_")] if rows else []
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        if cols:
            w.writeheader()
            w.writerows(rows)
    print(f"\nWrote {OUTPUT_CSV}")