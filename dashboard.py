"""
Dashboard for the email ranker (sponsorship or job-application mode).

    streamlit run dashboard.py
    (job mode:  set RANKER_MODE=job  first, or edit MODE in rank_companies.py)

Needs, in the same folder:
    rank_companies.py      (detection + ranking)
    emails_cache.json      (from fetch_threads.py or fetch_outlook.py)
    ranked_companies.csv   (from rank_companies.py)
"""

import json
import os

import pandas as pd
import streamlit as st

from rank_companies import PROFILE, SPONSOR_THRESHOLD as THRESHOLD, sponsor_score

YES = f"✅ {PROFILE['match_label']}"
THREAD = "🔗 Thread reply"
NO = f"❌ {PROFILE['no_label']}"

st.set_page_config(page_title="Email Ranker Dashboard", layout="wide")
st.title(f"{PROFILE['name'].title()} Dashboard")


@st.cache_data
def load_emails(path="emails_cache.json"):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


@st.cache_data
def load_ranking(path="ranked_companies.csv"):
    try:
        return pd.read_csv(path)
    except pd.errors.EmptyDataError:
        return pd.DataFrame()


tab_check, tab_rank = st.tabs(["Email check", "Company priorities"])

# ---------------------------------------------------------------------------
# Tab 1: does each email look relevant?
# ---------------------------------------------------------------------------
with tab_check:
    if not os.path.exists("emails_cache.json"):
        st.error("emails_cache.json not found. Run fetch_threads.py (or fetch_outlook.py) first.")
    else:
        emails = load_emails()
        if not emails:
            st.warning(
                "emails_cache.json has 0 emails, so there is nothing to show. "
                "The mailbox search probably matched nothing: widen the query "
                "(QUERY in fetch_threads.py, or days/max_emails in fetch_outlook.py) and re-run."
            )

        threshold = st.slider(
            "Score needed to count as a match (for testing here only)",
            min_value=1, max_value=10, value=THRESHOLD,
        )

        scored = [(e, *sponsor_score(e)) for e in emails]
        hit_threads = {e["thread_id"] for e, s, _ in scored
                       if e.get("thread_id") and s >= threshold}
        hit_domains = {e["domain"] for e, s, _ in scored
                       if not e.get("thread_id") and s >= threshold}

        rows = []
        for e, score, reasons in scored:
            if score >= threshold:
                verdict = YES
            elif (e.get("thread_id") in hit_threads) or (
                    not e.get("thread_id") and e["domain"] in hit_domains):
                verdict = THREAD
            else:
                verdict = NO
            rows.append({
                "Verdict": verdict,
                "Score": score,
                "Date": (e.get("date") or "")[:10],
                "From": e.get("sender_name") or e.get("sender_email", ""),
                "Domain": e.get("domain", ""),
                "Subject": e.get("subject", ""),
                "Why": "; ".join(reasons),
            })
        df = pd.DataFrame(rows, columns=["Verdict", "Score", "Date", "From", "Domain", "Subject", "Why"])

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Emails", len(df))
        c2.metric(PROFILE["match_label"], (df.Verdict == YES).sum())
        c3.metric("Thread replies", (df.Verdict == THREAD).sum())
        c4.metric(PROFILE["no_label"], (df.Verdict == NO).sum())

        show = st.multiselect(
            "Show", list(df.Verdict.unique()), default=list(df.Verdict.unique()))
        view = df[df.Verdict.isin(show)].sort_values("Score", ascending=False)
        st.dataframe(view, width="stretch", hide_index=True)

        st.subheader("Read one email")
        if len(view):
            labels = [f"{r.Date} | {r.From} | {r.Subject[:60]}" for r in view.itertuples()]
            pick = st.selectbox("Choose an email", range(len(labels)),
                                format_func=lambda i: labels[i])
            chosen = view.iloc[pick]
            match = next(e for e, _, _ in scored
                         if e.get("subject", "") == chosen.Subject
                         and e.get("domain", "") == chosen.Domain
                         and (e.get("date") or "")[:10] == chosen.Date)
            st.write(f"**{chosen.Verdict}** (score {chosen.Score}): {chosen.Why or 'no signals found'}")
            st.text_area("Body", match.get("body", ""), height=250)

# ---------------------------------------------------------------------------
# Tab 2: ranked companies (from rank_companies.py)
# ---------------------------------------------------------------------------
with tab_rank:
    if not os.path.exists("ranked_companies.csv"):
        st.error("ranked_companies.csv not found. Run rank_companies.py first.")
    else:
        ranked = load_ranking()
        if ranked.empty:
            st.info("ranked_companies.csv has no companies yet. Fetch some emails and run rank_companies.py again.")
            st.stop()
        intents = st.multiselect(
            "Intent", list(ranked.intent.unique()), default=list(ranked.intent.unique()))
        view = ranked[ranked.intent.isin(intents)]
        st.dataframe(view, width="stretch", hide_index=True)
        st.bar_chart(view.set_index("company")["final_score"])