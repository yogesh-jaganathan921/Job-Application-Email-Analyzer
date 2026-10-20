from sponsor_filter import filter_sponsor_emails

emails = filter_sponsor_emails(emails)
active, sponsored, rejected = rank(emails)