# WhatsApp Business API — Pricing Reference (Malaysia)

## 1. Free Tier (Meta Cloud API)

| Item | Detail |
|------|--------|
| Trial period | **90 days** free |
| Free conversations/month | 1,000 (ongoing, post-trial) |
| Max phone numbers (free test) | Up to **5 unique WhatsApp numbers** |
| After 5 users / trial ends | Must upgrade to paid Meta Business account |

---

## 2. Conversation Types & Rates (Malaysia)

| Type | Who starts it | Rate (MYR) | Window |
|------|--------------|-----------|--------|
| **User-initiated** | Customer sends first message ("Hi") | **RM 0.15** | 24 hrs |
| **Business-initiated** | You send first (broadcast/template) | **RM 0.45** | 24 hrs |
| **Utility / Auth** | System-triggered (OTP, order confirm) | **RM 0.08** | 24 hrs |

All messages in one 24-hour window = **1 conversation fee**, regardless of message count.

---

## 3. Your Bot Flow (User-Initiated = RM 0.15)

User sends "Hi" → bot auto-replies → session continues.
Still billed as **user-initiated (RM 0.15)** because the customer sent the first message.
You are only charged RM 0.45 if YOUR bot sends the FIRST message (broadcast).

---

## 4. Scaling Cost Table — User-Initiated

| Daily active users | Per day | Per month (30d) | Per year |
|--------------------|---------|-----------------|----------|
| 5                  | RM 0.75 | **RM 22.50**    | RM 270   |
| 10                 | RM 1.50 | **RM 45**       | RM 540   |
| 20                 | RM 3.00 | **RM 90**       | RM 1,080 |
| 50                 | RM 7.50 | **RM 225**      | RM 2,700 |
| 100                | RM 15   | **RM 450**      | RM 5,400 |
| 500                | RM 75   | **RM 2,250**    | RM 27,000|

---

## 5. Broadcast Pricing (Business-Initiated = RM 0.45)

Requires a **Meta-approved Message Template** (approval takes 1–3 days).
If user replies within 24 hrs, that window is still billed at RM 0.45 (not RM 0.15).

| Daily broadcast recipients | Per day  | Per month |
|---------------------------|----------|-----------|
| 10                        | RM 4.50  | RM 135    |
| 50                        | RM 22.50 | RM 675    |
| 100                       | RM 45    | RM 1,350  |
| 500                       | RM 225   | RM 6,750  |

---

## 6. Hosting Cost (to run bot 24/7)

| Platform | Cost/month | Notes |
|----------|------------|-------|
| **Railway** | ~USD 5–20 (~RM 23–94) | Easy deploy, auto-scale |
| **Render** | Free / USD 7 (~RM 33) | Free tier sleeps after inactivity |
| **VPS (DigitalOcean / Contabo)** | USD 4–12 (~RM 19–56) | Always on, full control |

**Recommended: Railway (~RM 30–50/month) or a cheap VPS.**

---

## 7. Total Monthly Budget Estimates

| Scenario | WhatsApp | Hosting | Total/month |
|----------|----------|---------|-------------|
| 5 users/day, no broadcast | RM 22.50 | RM 30 | **~RM 53** |
| 20 users/day, no broadcast | RM 90 | RM 30 | **~RM 120** |
| 50 users/day, no broadcast | RM 225 | RM 30 | **~RM 255** |
| 100 users/day, no broadcast | RM 450 | RM 50 | **~RM 500** |
| 50 users/day + 100 broadcasts/day | RM 225+RM1,350 | RM 50 | **~RM 1,625** |

---

## 8. Setup Steps

1. Create **Meta Business Account** at business.facebook.com
2. Add **WhatsApp** product in Meta Developer Console
3. Register a phone number (new SIM or virtual number)
4. Deploy your bot to Railway / VPS → get a public HTTPS URL
5. Set the URL as the **webhook** in Meta Developer Console
6. For broadcasts: submit **Message Templates** for Meta review

---

*Malaysia rates verified from community sources, May 2026.*
*Always check current rates: developers.facebook.com/docs/whatsapp/pricing*
