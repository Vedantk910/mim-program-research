"""Build MIM_TOP25_AND_MORE.xlsx — Top 25 + 15 near-miss programs for Vedant Konde.
Data source: MIM_TOP25_REPORT.md (6 regional research passes + official-page verifications).
Rows marked Confirmed = verified against official pages this research cycle; Est = 2024-25 vintage, re-check on 2027 intake pages before paying fees.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = r"C:\Finance Projects\MIM_TOP25_AND_MORE.xlsx"

HEADERS = [
    "RANK", "PROGRAM", "UNIVERSITY", "COUNTRY", "TIER", "DURATION", "2027 DEADLINE",
    "ODDS @ GMAT FOCUS 655", "TUITION (2027 EST.)", "ALL-IN COST (TUITION+LIVING)",
    "% INDIAN (EST.)", "STARTING SALARY", "ROI: RECLAIM ALL COST",
    "WORK WHILE STUDYING", "POST-STUDY VISA", "SPOON-FED /10", "EAS INDIAN LIVING",
    "DATA CONFIDENCE", "VERDICT (1-LINE)",
]

# rank, program, uni, country, tier, duration, deadline, odds, tuition, allin, pind, salary, roi, work, visa, spoon, living, confidence, verdict
TOP25 = [
    (1, "Master in Management (MiM)", "NUS (Singapore)", "SG", "T1", "1 yr",
     "~mid-Feb 2027 (single round; unverified)", "HIGH",
     "S$56,700 pre-GST CONFIRMED (AY24/25); 2027 est S$63-66k", "S$85-90k (~INR 54-58L)",
     "30-40%", "S$55-65k", "~2.2 yrs",
     "PWP: 16h/wk term, 24h/wk breaks (earns S$15-18k/yr)", "EP + COMPASS (NUS = bonus institution)",
     "7.5", "9/10", "Tuition CONFIRMED; salary/deadline est.",
     "Flagship for an Indian: IIM-A brand signal, huge Indian community"),
    (2, "MIM (Business School)", "Dublin City University", "IE", "T1", "12-18 mo",
     "Rolling ~Apr-Jun 2027 (est)", "HIGH (GMAT not required; work exp valued)",
     "~EUR 33-38k non-EU (est)", "EUR 55-65k (~INR 48-57L)", "20-30%",
     "EUR 45-55k", "~2-2.5 yrs — best absolute ROI",
     "FULL-TIME allowed (EU, no 20h cap)", "Stamp 1G up to 24 months — VERIFIED",
     "5-7", "Good", "Visa CONFIRMED; tuition est.",
     "Best value: cheapest Tier-1 + full work rights + Dublin IB/tech hub"),
    (3, "Master of Management (MiM)", "HKUST", "HK", "T1", "1 yr",
     "Rolling ~Mar 15 2027 (unverified)", "MED-HIGH (GMAT optional)",
     "HK$385k (2025; 2027 est HK$400-420k)", "HK$550k (~INR 46L) — cheapest APAC",
     "25-35%", "HK$420-480k", "~1.5 yrs",
     "FULL-TIME from month 7 (best earner on list)", "IANG 2-yr post-study stay",
     "7", "7.5/10", "Est.",
     "Smart APAC on a budget: ~30% cheaper than SG trio in INR"),
    (4, "Master in Mgmt in Finance (MiF)", "SMU", "SG", "T1", "10 mo",
     "~Mar 31 2027 (unverified)", "MED-HIGH (class avg ~670-700)",
     "S$72-78k (est)", "S$95-100k (~INR 61-64L)", "30-35%",
     "S$65-75k (highest SG trio)", "~1.7 yrs",
     "PWP 16-24h/wk", "EP + COMPASS (SMU = bonus institution)",
     "8.5 (HIGHEST)", "9/10", "Est.",
     "Mandatory Practicum = closest thing to IIM placements in APAC"),
    (5, "Master in Management (finance electives)", "TUM Munich", "DE", "T2", "3 yrs",
     "~Jul 2027 (confirm mim.tum.de)", "MED",
     "NEAR-ZERO (Bavaria abolished tuition; ~EUR 350-550/sem admin)", "~EUR 45-50k (living only)",
     "5-8%", "EUR 55-60k", "~2.5-3.5 yrs",
     "140 full days/yr (~1-2 days/wk)", "18-month job-search residence — VERIFIED",
     "6", "Good (Munich)", "Public fee policy CONFIRMED; salary/deadline est.",
     "Best risk-adjusted DACH: top brand + free tuition + 18-mo safety net"),
    (6, "Master in Management", "WHU (Oftersheim)", "DE", "T2", "21 mo",
     "Rolling; RESPONSIBLE LEADER SCHOLARSHIP deadline 31 Jan (est)", "MED (exp compensates; class ~700)",
     "~EUR 50k; 75% 'Responsible Leader' award (CONFIRMED) -> ~EUR 20-25k",
     "~EUR 75k -> ~EUR 20-25k w/ scholarship", "3-5%",
     "EUR 67,149 avg (PUBLISHED on WHU page)", "<6 mo post-grad w/ scholarship",
     "140 days/yr", "18-month job-search", "6.5", "Small town; Heidelberg 15 min",
     "Scholarship + avg salary CONFIRMED; rest est.",
     "Highest-leverage scholarship action on the whole list — apply by 31 Jan"),
    (7, "MIM (MSc Management & Intl Mgmt)", "Frankfurt School", "DE", "T2", "1 yr",
     "~Mar 15 2027 (unverified)", "MED-HIGH",
     "~EUR 38-43k (est)", "EUR 55-60k", "5-7%", "EUR 58-65k",
     "~1.5-2 yrs", "140 days/yr", "18-month job-search",
     "7.5", "Good", "Est.",
     "Most spoon-fed German banking pipeline — Deutsche Bank HQ town"),
    (8, "Master in Management (MiM)", "Imperial College London", "UK", "T2", "12 mo",
     "6 Jan 2027 (CONFIRMED)", "MED (0-5 yrs career-change friendly)",
     "GBP 51,000 CONFIRMED", "GBP 70-72k (London premium)", "30-40% Asian",
     "GBP 48-52k; 94% placed within 6 mo (CONFIRMED)", "~2.5-3.5 yrs",
     "20h/wk term + unlimited in vacation", "2-yr Graduate Route (confirmed active for 2027)",
     "7", "9/10 (London)", "Tuition + placement CONFIRMED",
     "Best 5-star UK MiM; skews consulting, not pure IB"),
    (9, "MSc in Finance", "SSE (Handelshogskolan Stockholm)", "SE", "T2", "2 yrs",
     "Early ~Nov 15 / General ~Jan 15 (2027 unverified)",
     "HIGH (verified GMAT floor 555 new-format)",
     "SEK 360,000 CONFIRMED (~EUR 32.7k)", "EUR 60-68k", "not published",
     "SEK 550-650k (est)", "~3.5-4 yrs part-time", "20h/wk",
     "~12-month job-search (some sources say 18 — verify)", "6",
     "Clean, safe, cold long winter", "Tuition CONFIRMED; salary est.",
     "Deepest Nordic IB/PE pipeline (EQT network)"),
    (10, "Master in Management", "NTU Nanyang", "SG", "T2", "1 yr",
     "Jan round ~Jan 31 2027 (unverified)", "HIGH (most generous GMAT stance)",
     "~S$69k (est)", "S$90-95k (~INR 58-61L)", "35-45% (most India-heavy SG)",
     "S$55-65k", "~2.1 yrs", "PWP 16-24h/wk", "EP + COMPASS",
     "7", "9/10", "Est.",
     "Safest SG admission + mgmt+analytics dual skill"),
    (11, "MSc in Finance", "HKU", "HK", "T2", "1 yr",
     "~Mar 31 2027 (unverified)", "MED — program avg ~700; admission risk item",
     "HK$390-410k (est)", "HK$560k (~INR 47L)", "20-30% (cohort ~150)",
     "HK$500-600k — highest APAC", "~1.3 yrs", "Full-time from month 7",
     "IANG 2-yr", "6.5", "7/10", "Est.",
     "Reach-quant: best APAC salary, but 655 is below the admit band"),
    (12, "MSc in Management (MiM)", "Warwick Business School", "UK", "T3", "1 yr",
     "Rolling, last round ~Feb 2027 (est)", "MED",
     "~GBP 52-55k (2025/26; 2027 unconfirmed)", "GBP 70-75k", "30-40% Asian",
     "GBP 45-50k", "~2.5-3.5 yrs", "20h/wk + unlimited vacation",
     "2-yr Graduate Route", "7.5", "Good (Birmingham cheaper)", "Est.",
     "Best UK value MiM — consulting placement machine, cheap city"),
    (13, "MSc in Finance", "KTH Royal Institute", "SE", "T3", "2 yrs",
     "~Mar 1 2027 (est)", "MED-HIGH",
     "~SEK 385k (~EUR 35k); KTH scholarship covers 1-2 yrs tuition (CONFIRMED range)",
     "EUR 60-66k (scholarship halves)", "not published",
     "SEK 550-650k (est)", "~3 yrs (faster w/ scholarship)", "20h/wk",
     "~12-month job-search", "5.5", "Good", "Scholarship tiers CONFIRMED; fee est.",
     "Quant-leaning; scholarship = fast ROI"),
    (14, "MSc in Finance", "LSE", "UK", "T3", "9 mo",
     "Single early round ~mid-Jan 2027 (unconfirmed)", "LOW-MED — REACH (admit band 660-700+)",
     "~GBP 49,974 (2025/26; 2027 unconfirmed)", "GBP 68-72k",
     "~40-50% Asian (forum est)", "GBP 70-80k (IB analyst)",
     "~1.5-2.5 yrs IF admitted", "20h/wk + vacation", "2-yr Graduate Route",
     "6", "9/10", "2025/26 tuition published; rest est.",
     "Best UK IB pipeline — apply as the reach"),
    (15, "MiM in Finance (1-yr track)", "ESSEC", "FR", "T3", "1 yr",
     "Rolling; early rounds ~Dec 2026 (est)", "MED (class avg ~670-690)",
     "~EUR 30-33k 1-yr (est); 2-yr Intensive EUR 40k+ (3rd-party verified)",
     "EUR 45-48k", "10-15%",
     "start ~EUR 65k; 3-yr avg EUR 80-90k (est)", "~1.5 yrs incl. PT work",
     "964 h/yr (≈48h/wk legal, France non-EU cap)", "APS 12-mo", "7", "Good",
     "Tuition est.; check apprenticeship (employer-paid) track on ESSEC page",
     "Big-bank recruiting + the best PT-work regime in France"),
    (16, "Master in Management (Grande Ecole)", "HEC Paris", "FR", "T3", "2 yrs",
     "ROUND 2 = 26 Nov 2026 (CONFIRMED, hec.edu)",
     "LOW-MED (class GMAT ~690-710; 655 is 30-45 pts below)",
     "EUR 61,655 CONFIRMED (incl. non-EU surcharge + app fee)",
     "EUR 95-100k", "15-20% (of ~550, est)",
     "EUR 118.5k AVG @3 YRS (Financial Times, VERIFIED); start EUR 70-85k",
     "~1.5-2 yrs incl. PT work", "964 h/yr (≈48h/wk)",
     "APS 12-mo -> Passeport Talent", "7", "Good",
     "Tuition + 3-yr salary CONFIRMED; % Indian + odds est.",
     "Top EU MiM brand; FT-verified salary fixes the ROI; 655 is the risk"),
    (17, "MSc in Finance (MFA)", "ETH Zurich", "CH", "T3", "2 yrs",
     "~15 May 2027 (unverified)", "MED (degree-fit driven, GMAT not required)",
     "~CHF 10-11k TOTAL tuition (est — near-zero)", "CHF 50-55k (living)",
     "10-20% (cohort ~60-100, est)", "CHF 95-120k (est)",
     "~1-1.5 yrs", "0 HOURS (CH: non-EU students cannot work during study)",
     "NONE — L permit, employer-tied, ~CHF 70k+ salary floor", "5.5",
     "Cold, very expensive", "Tuition est.; CH work/visa rules confirmed pattern",
     "Quant-elite + near-zero tuition; land the offer or leave"),
    (18, "Master in Mgmt of Finance", "St. Gallen SFA", "CH", "T4", "1 yr",
     "~May-Jun 2027 (unverified)", "MED-HIGH (class avg ~620-670)",
     "~CHF 42-45k (agent est) vs CHF 90-100k (2-yr variant) — RE-VERIFY ON SFA PAGE",
     "CHF 48-52k", "15-25% (est)", "CHF 90-105k @3 yrs (est)",
     "~1-1.5 yrs", "0 HOURS (CH)", "NONE — L permit employer-tied",
     "7 (most spoon-fed Swiss route)", "Alpine cold, small town", "Est.; CONFLICTING fee figures",
     "Direct UBS/Pictet/Julius Baer pipeline; German for long-term life"),
    (19, "MSc Banking & Finance", "UZH (Zurich)", "CH", "T3", "2 yrs",
     "~31 Mar 2027 (unverified)", "MED-HIGH (exam-based, no GMAT)",
     "~CHF 1,800-1,900/sem -> ~CHF 7-8k TOTAL (est)", "CHF 55-65k (living)",
     "3-5% (est)", "CHF 95-105k — highest DACH", "~1.5 yrs",
     "0 HOURS (CH)", "NONE — B permit via employer", "5", "Cold",
     "Est.", "Exam-based admission + top salary; no safety net"),
    (20, "Master in Finance & Accounting", "WU Vienna", "AT", "T4", "2 yrs",
     "~31 Mar 2027 (unverified)", "HIGH (exam-based, no GMAT)",
     "~EUR 15k/yr non-EU (est)", "EUR 55-60k", "2-4% (est)",
     "EUR 55-60k", "~3.5 yrs", "~18h/wk (est)",
     "6-month job-search (shortest EU)", "5", "BEST in DACH for Indians",
     "Est.", "Easiest admission + best Indian living; grads move to FRA/ZRH"),
    (21, "MSc in Financial Economics (RENAMED from Finance & Insurance)", "Maastricht", "NL", "T4",
     "2 yrs", "~1 Jun 2027 non-EU (agent est)", "HIGH",
     "EUR 21,360/yr non-EU CONFIRMED (~EUR 43k total)", "EUR 43-45k",
     "not published", "EUR 45-55k (est)", "~2-2.5 yrs",
     "16h/wk", "12-mo zoekjaar, RENEWABLE; apply within 3 yrs of graduation (VERIFIED)",
     "6", "Good", "Tuition CONFIRMED; deadline est.",
     "Cheap NL + renewable search window; deadline ~Jun 2027"),
    (22, "MSc Banking & Finance", "CBS Copenhagen", "DK", "T4", "2 yrs",
     "~Jan 31 2027 (est)", "HIGH",
     "EUR 16k/yr non-EU CONFIRMED (~EUR 32k total)", "EUR 32-35k",
     "not published", "EUR 50-60k (est)", "~2.5 yrs",
     "20h/wk", "12-mo post-study + 1-yr extension (verified)", "6.5",
     "Good — most Indian-friendly Nordic city", "Tuition CONFIRMED; salary est.",
     "Cheapest VERIFIED EU finance degree; good consulting/IB tilt"),
    (23, "MSc in Finance (SDA Bocconi)", "Bocconi Milan", "IT", "T4", "2 yrs",
     "Single deadline ~late Feb-Mar 2027 (est)", "LOW-MED (class ~700; 655 -> entrance exam likely)",
     "~EUR 44-47k/yr [2024-25]", "EUR 115-125k", "15-20% (est)",
     "EUR 55-68k [2024-25]", "~2.5-3.5 yrs", "20h/wk (Italy cap)",
     "12-mo job-search (verify)", "7", "Good", "[2024-25] vintage, unverified",
     "Best EU quant track (fixed bank roster: Intesa, UniCredit, UBS/JPM Milan); cost = 2x IE"),
    (24, "Master in Management (MIM)", "IE Business School, Madrid", "ES", "T3",
     "15 mo (10 mo classes + 5 mo internship)",
     "R1 15 Nov 2026 / R2 15 Feb 2027 / R3 15 May 2027 (CONFIRMED, ie.edu)",
     "MED-HIGH (class avg ~660 FLEX CONFIRMED; 655 FOCUS near/above adjusted)",
     "EUR 51,200 CONFIRMED (incl. IE Foundation)", "EUR 72-78k",
     "10-15% — LARGEST INDIAN COMMUNITY IN EUROPE",
     "~EUR 50-55k start; US$95,000 3-yr weighted (CONFIRMED); 88% employed in 3 mo (CONFIRMED)",
     "~0.8-1.5 yrs — FASTEST EU ROI (R1 merit award EUR 15-35k est)",
     "Up to 30h/wk (Spanish non-EU rule 2023+; confirm with consulate)",
     "12-mo 'permiso de permanencia' (+6-mo renewal if starting a company)",
     "6", "BEST — 500+ Indian restaurants, India Fest",
     "Tuition/rounds/placement CONFIRMED; salary EUR median + work-rule est.",
     "EU all-rounder #1: fastest ROI + biggest Indian support network in Europe"),
    (25, "MSc in Finance", "Ca' Foscari Venice", "IT", "T4", "2 yrs",
     "Single national call ~mid-Nov-mid-Dec 2026 (est)", "MED-HIGH (GMAT optional)",
     "~EUR 3,500-4,500/yr public non-EU cap [2024-25]",
     "EUR 25-30k", "3-6% (est)", "EUR 35-45k [2024-25]",
     "~8-12 MONTHS — best raw ROI on the entire list", "20h/wk (Italy cap)",
     "12-mo job-search", "4", "Moderate (smallest community)",
     "[2024-25] vintage, unverified",
     "Cheap quant foundation; self-market; launchpad to London/Zurich in year 3"),
]

# Near-miss 15 (ranks 26-40)
NEARMISS = [
    (26, "Master in Management (career-change)", "Oxford SAID", "UK", "T4", "9 mo",
     "~mid-Jan + late-Feb 2027 (est)", "LOW-MED",
     "~GBP 62-65k (2025/26; 2027 unconfirmed)", "GBP 80-85k", "30-40% Asian",
     "GBP 50-58k (est)", "~3.5-4.5 yrs — worst UK ROI", "20h/wk + vacation",
     "2-yr Graduate Route", "6", "Good", "Est.",
     "Brand-only play; no standard MiM merit scholarship; diluted career-switcher cohort"),
    (27, "MSc in Finance", "Cambridge", "UK", "T4", "1 yr",
     "~early Jan 2027 (est)", "LOW (very selective quant brand)",
     "~GBP 45-52k [2024-25]", "GBP 75-80k", "not published",
     "GBP 70-75k (est)", "~2.5-3.5 yrs", "20h/wk", "2-yr Graduate Route",
     "6", "Good", "Est.",
     "Prestige quant degree; 655 Focus is below their typical band — reach"),
    (28, "Master in Management", "EDHEC", "FR", "T4", "12 mo + 4-mo internship",
     "Rounds; early ~Dec 2026 (est)", "MED-HIGH (your 655 is AT class median ~650-670)",
     "EUR 51,250 non-EU CONFIRMED (EU: 44,700)", "~EUR 60k", "~10% (est)",
     "EUR 50-55k (est)", "~2-2.5 yrs", "964 h/yr (France)", "APS 12-mo",
     "6.5", "Good (cheap Lille living)", "Tuition CONFIRMED; rest est.",
     "Best cost-performance Grande Ecole; a notch below HEC/ESSEC on brand"),
    (29, "Master in Management (EPFL Lausanne)", "NEOMAP", "CH", "T4", "1.5 yrs",
     "~Mar 2027 (est)", "MED",
     "~CHF 46-48k (est)", "CHF 65-70k", "est",
     "0 HOURS (CH)", "NONE (L permit employer-tied)", "5", "Cold",
     "Est.",
     "Dominate-by-ETH: same CH 0-work rule + weaker brand — skip unless EPFL research is the goal"),
    (30, "MSc in Finance", "Swiss Finance Institute (Zurich)", "CH", "T4", "3 sems + project",
     "~Mar 2027 (est)", "MED-HIGH",
     "~CHF 36-40k (est)", "CHF 50-55k", "est",
     "CHF 85-95k (est)", "~1.5-2 yrs", "0 HOURS (CH)",
     "NONE (L permit employer-tied)", "6", "Cold", "Est.",
     "Known among Indian quant shops; pure self-marketing; verify fees on sfi.edu"),
    (31, "MSc in Finance", "GSE (Chalmers + Univ. Gothenburg)", "SE", "T4", "2 yrs",
     "~Mar 1 2027 (est)", "HIGH (GMAT not strictly required)",
     "~SEK 450-500k for 2 yrs [est]", "EUR 55-65k", "not published",
     "SEK 500-600k (est)", "~2.5-3 yrs", "20h/wk",
     "~12-month job-search", "5.5", "Good (cheaper CoL than Stockholm)",
     "Est.",
     "SEB/Nordea/Handelsbanken quant teams sit in Gothenburg; strong quant value"),
    (32, "MSc in Finance", "Lund University", "SE", "T4", "2 yrs",
     "~Mar 1 2027 (est)", "HIGH (no GMAT requirement)",
     "~SEK 150-200k/yr [est]", "EUR 55-65k", "not published",
     "SEK 500-580k (est)", "~2.5-3 yrs", "20h/wk",
     "~12-month job-search", "5.5", "Good (Malmo = SE's most multicultural city)",
     "Est.",
     "Cheapest big-name SE finance degree; value fallback"),
    (33, "Master in Management (MIM)", "Aalto University", "FI", "T4", "2 yrs",
     "~Mar 1 2027 (est)", "HIGH (GMAT optional)",
     "~EUR 15-20k/yr [est]", "EUR 30-40k", "not published",
     "EUR 50-60k (est)", "~2.5-3 yrs", "20h/wk",
     "6-month job-search (shortest here)", "5", "Dark winter (Dec-Jan)",
     "Est.",
     "Cheapest big-name in Nordics, but 6-mo search + weakest finance ecosystem"),
    (34, "MSc in Finance", "LMU Munich", "DE", "T4", "2 yrs",
     "~15 Jul 2027 (est)", "MED-HIGH",
     "Near-zero (Bavarian public fee ~EUR 350/sem)", "~EUR 30-35k (living only)",
     "3-5% (est)", "EUR 55-60k (est)", "~3-3.5 yrs",
     "140 days/yr", "18-month job-search", "5", "Good",
     "Est.",
     "Cheap public option; self-marketed placements (weaker than TUM/FS)"),
    (35, "MiM", "EBS Business School", "DE", "T4", "1.5 yrs",
     "~Mar 2027 (est)", "HIGH",
     "~EUR 20-25k (est)", "EUR 35-40k", "2-4% (est)",
     "EUR 45-50k (est)", "~2-2.5 yrs", "140 days/yr",
     "18-month job-search", "5.5", "Good (Rhein-Main/Frankfurt access)",
     "Est.",
     "Smaller brand; decent Frankfurt-banking access at low cost"),
    (36, "MSc in Financial Mathematics (quant variant)", "Erasmus Rotterdam", "NL",
     "T4", "1 yr", "~Feb-Mar 2027 (est)", "MED-HIGH (class avg ~650-680)",
     "~EUR 14-16k non-EU (est)", "EUR 28-30k", "10-20% (est)",
     "EUR 50-65k (est)", "~1-1.5 yrs incl. PT work",
     "16h/wk (+40h/wk summer)", "12-mo zoekjaar, RENEWABLE (verified)",
     "6.5", "Good", "Est.",
     "CHEAPEST QUANT FEEDER on the list — the Financial Mathematics variant is the real quant route"),
    (37, "MSc in Finance", "Tilburg University", "NL", "T4", "1 yr",
     "~Feb-Mar 2027 (est)", "HIGH",
     "~EUR 12-14k (est)", "EUR 25-28k", "est",
     "EUR 40-50k (est)", "~2-2.5 yrs", "16h/wk",
     "12-mo zoekjaar (renewable)", "5", "Good", "Est.",
     "Weakest quant brand of the NL trio; placements skew consulting/accounting"),
    (38, "MSc in Finance", "ESADE", "ES", "T4", "12 mo",
     "R1 ~15 Dec 2026 (est)", "MED (class ~650-700)",
     "~EUR 45-48k [2024-25]", "EUR 65-70k", "5-10% (est)",
     "EUR 50-55k (est)", "~1.2-1.8 yrs", "Up to 30h/wk (ES rule)",
     "12-mo 'permiso de permanencia'", "6", "Good (Barcelona)",
     "[2024-25] vintage, unverified",
     "Decent middle: smaller Indian community than IE, quant track thin — drop unless Barcelona lifestyle decides it"),
    (39, "MSc in Financial Management & Business Analysis", "CEU (Vienna/Budapest)", "AT/HU",
     "T4", "18-24 mo", "~Dec 2026 (est)", "HIGH",
     "~EUR 10-12k [est]", "EUR 15-20k", "est",
     "EUR 35-45k (est)", "~1-1.5 yrs", "30h/wk (EU)",
     "6-month job-search (AT/HU)", "4.5", "Vienna/Budapest (moderate community)",
     "Est.",
     "Low-cost EU brand + low macro risk; weakest quant ceiling of the lot"),
    (40, "MSc in Finance (EXISTS? — VERIFY)", "NYU Abu Dhabi", "AE", "T4", "1 yr?",
     "~Dec 2026-Jan 2027 (est)", "MED (if program exists)",
     "UNCONFIRMED — NYUAD historically covers tuition via institutional scholarships; standard rate would be US$60k+/yr",
     "UNCONFIRMED", "very small cohort",
     "thin local quant market (ADGM)", "only good if funding confirmed",
     "NONE — UAE student visa allows no work",
     "6-mo 'seeker of opportunities' (extendable to 12m) + Golden Visa options",
     "3", "Easy (huge Indian diaspora; heat)",
     "UNCONFIRMED — program existence + post-2028 delivery model both unverified",
     "SKIP unless: (a) finance MSc confirmed at nyuad.edu, (b) funding confirmed, (c) you accept 0 work rights while studying"),
]

CONFIDENCE_LEGEND = [
    ("CONFIRMED", "Verified against official program/university pages this research cycle (URLs in MIM_TOP25_REPORT.md). Safe to plan around."),
    ("EST / [unverified]", "2024-25 vintage or agent estimate; search backend was throttled. Re-check on the official 2027 intake page BEFORE paying any application fee."),
    ("[2024-25]", "Explicitly last-year published figures carried over; expect +5-8% for 2027."),
    ("VERIFY", "Fact could not be confirmed by any source this cycle (e.g., NYUAD finance MSc existence, St. Gallen 1-yr fee)."),
]

STRATEGY = [
    ("APPLICATION WAVES (Dec 2026 - Mar 2027)", ""),
    ("Wave 1 (Nov-Dec 2026):", "HEC R2 (26 Nov 2026 - CONFIRMED), IE R1 (15 Nov 2026 - CONFIRMED), ESSEC rolling, WHU scholarship round (31 Jan — apply early)"),
    ("Wave 2 (Jan 2027):", "Imperial (6 Jan - CONFIRMED), NTU Jan round (~31 Jan), LSE single round (~mid-Jan), Cambridge (~early Jan), WHU general (~31 Jan)"),
    ("Wave 3 (Feb-Mar 2027):", "SMU (~Mar 31), HKUST (~Mar 15), HKU (~Mar 31), Bocconi (~late Feb-Mar), Erasmus/Tilburg (~Feb-Mar), Maastricht (~Jun - non-EU window)"),
    ("Wave 4 (Apr-Jul 2027):", "Ca' Foscari single call, CEU, DCU rolling, TUM (~Jul), WU/UZH/ETH/St. Gallen (~Mar-May)"),
    ("ANCHOR SET (Tier 1 = 5 + 1)", ""),
    ("Apply first:", "NUS, DCU, HKUST, SMU, TUM + IE Madrid (EU all-rounder)"),
    ("Reach layer:", "HKU MScF, LSE, HEC, Cambridge"),
    ("Safety layer:", "NTU, Warwick, DCU, WU, Aalto, Erasmus, Tilburg, Ca' Foscari"),
    ("SCHOLARSHIP LEVERAGE (apply early!)", ""),
    ("WHU Responsible Leader 75% (31 Jan deadline, CONFIRMED)", "Single highest-leverage application on the whole list — cuts EUR 75k to ~EUR 20-25k"),
    ("KTH 1-2 yr tuition coverage", "Verified tiered scholarship from KTH admissions page"),
    ("IE R1 merit (EUR 15-35k est.)", "Early round = best odds"),
    ("NUS/NTU/SMU merit awards", "S$10-30k band; no India-specific MIM scholarships in SG"),
    ("ESSEC apprenticeship track", "Employer PAYS tuition + salary while studying — verify which track it covers before 2027"),
    ("RECLAIM-MONEY-WHILE-STUDYING LADDER", ""),
    ("1st: HK (full-time from month 7)", "~INR 10-15L back in final semester alone"),
    ("2nd: Ireland (full-time all year)", "Best for staying in EU"),
    ("3rd: France (964 h/yr ≈ 48h/wk)", "Best PT regime in continental EU"),
    ("4th", "Spain (30h/wk) / SG PWP (16-24h/wk) / UK+NL+Nordics (16-20h/wk) / Germany (140 days/yr)"),
    ("0th: Switzerland & UAE (0 hours)", "Cannot work at all during study"),
    ("POST-STUDY JOB-SEARCH LADDER (longest first)", ""),
    ("Ireland 24-mo", "Stamp 1G, verified"),
    ("UK 24-mo", "Graduate Route"),
    ("Germany 18-mo", "Aufenthaltstitel zur Jobsuche"),
    ("France 12-mo", "APS -> then Passeport Talent"),
    ("Netherlands 12-mo", "zoekjaar, RENEWABLE, apply within 3 yrs"),
    ("Spain / Sweden / Denmark 12-mo", "Denmark adds 1-yr extension"),
    ("Austria / Finland 6-mo", "Shortest"),
    ("Switzerland / UAE: none", "Must already hold an offer"),
    ("BEFORE YOU PAY ANY FEE (re-verify list)", ""),
    ("1", "NYU Abu Dhabi: does a finance MSc even exist + post-2028 delivery model (nyuad.edu)"),
    ("2", "St. Gallen SFA: 1-yr fee — CHF 42-45k vs CHF 90-100k conflict (sfa-egi.ch)"),
    ("3", "ESSEC: 1-yr MiF tuition vs 2-yr Intensive track (essec.edu)"),
    ("4", "Bocconi: 2027 fee + entrance exam requirement at 655 Focus (sdabocconi.it)"),
    ("5", "Spain 30h/wk work rule with your consulate (2023 consolidated text)"),
    ("6", "Sweden 12 vs 18-month job-search permit (Migrationsverket)"),
    ("7", "UK 2027 Skilled Worker salary threshold + Graduate Route still 2 yrs"),
    ("8", "SG EP/COMPASS 2027 salary floor (mom.gov.sg) and HK IANG rules (immd.gov.hk)"),
]

def style_sheet(ws, widths, freeze="C2"):
    thin = Side(style="thin", color="D0D0D0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    hdr_fill = PatternFill("solid", fgColor="1D1D1F")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    tier_fill = {
        "T1": PatternFill("solid", fgColor="C6EFCE"),
        "T2": PatternFill("solid", fgColor="DDEBDD"),
        "T3": PatternFill("solid", fgColor="FFF2CC"),
        "T4": PatternFill("solid", fgColor="FCE4D6"),
    }
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.fill = hdr_fill
        c.font = hdr_font
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = border
    ws.row_dimensions[1].height = 42
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = border
            c.font = Font(size=10)
        tier = str(row[4].value) if row[4].value else ""
        if tier in tier_fill:
            row[4].fill = tier_fill[tier]
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions

wb = Workbook()

# Sheet 1: Top 25
ws = wb.active
ws.title = "Top 25 Ranked"
ws.append(HEADERS)
for r in TOP25:
    ws.append(list(r))
style_sheet(ws, [5, 26, 22, 7, 6, 9, 24, 22, 26, 20, 14, 20, 16, 24, 24, 9, 14, 30, 34])

# Sheet 2: Near-misses
ws2 = wb.create_sheet("15 Near-Misses (26-40)")
ws2.append(HEADERS)
for r in NEARMISS:
    ws2.append(list(r))
style_sheet(ws2, [5, 26, 22, 7, 6, 9, 24, 22, 26, 20, 14, 20, 16, 24, 24, 9, 14, 30, 34])

# Sheet 3: Strategy
ws3 = wb.create_sheet("Strategy & Deadlines")
ws3.append(["SECTION", "DETAIL"])
for a, b in STRATEGY:
    ws3.append([a, b])
ws3.column_dimensions["A"].width = 46
ws3.column_dimensions["B"].width = 110
for c in ws3[1]:
    c.fill = PatternFill("solid", fgColor="1D1D1F")
    c.font = Font(bold=True, color="FFFFFF")
for row in ws3.iter_rows(min_row=2):
    row[0].font = Font(bold=True, size=10)
    row[1].font = Font(size=10)
    row[1].alignment = Alignment(wrap_text=True, vertical="top")
    row[0].alignment = Alignment(vertical="top")

# Sheet 4: Read-me
ws4 = wb.create_sheet("Read-Me & Flags")
ws4.append(["TOPIC", "NOTE"])
readme = [
    ("CANDIDATE", "Vedant Konde — Indian, 21, B.E. Comp Eng 8.5/10, GMAT FOCUS 655, 2 yrs Amazon Risk Specialist (top-5% SME), chess-venture founder, AWS ML+DE certs. Sep 2027 intake. India/US/Australia excluded."),
    ("FX CONVERSIONS USED", "1 SGD ~ INR 6.4; 1 HKD ~ INR 0.84; 1 EUR ~ INR 98; 1 CHF ~ INR 88; 1 GBP ~ INR 160; 1 SEK ~ EUR 0.09. Update before financial planning."),
    ("CONFIDENCE LEGEND", ""),
]
for a, b in CONFIDENCE_LEGEND:
    readme.append((a, b))
readme += [
    ("SOURCES", "Confirmed-live: NUS MiM fees (mim.nus.edu.sg), Imperial (imperial.ac.uk), WHU scholarship+avg salary (whu.edu), HEC fees/rounds/salary (hec.edu, FT 2026), ESSEC 2-yr (3rd-party), IE everything (ie.edu), Maastricht fee (maastrichtuniversity.nl), CBS fee (cbs.dk), EDHEC fee (edhec.edu), KTH fee+scholarships (kth.se), SSE fee+GMAT floor (hhs.se), Ireland Stamp 1G (irishimmigration.ie + TCD), Germany 18-mo (Berliner Welcome Center), Sweden 12-mo (Migrationsverket), NL 3-yr renewal (UvA/IND)."),
    ("KNOWN CONFLICTS", "St. Gallen 1-yr fee: CHF 42-45k vs CHF 90-100k (different programs) — verify. Sweden job-search: 12 vs 18 months — verify."),
    ("OPEN ITEMS", "NYUAD finance MSc existence; ESSEC 1-yr vs 2-yr track cost; Spain 30h/wk consulate confirmation; UK 2027 Skilled Worker threshold; SG 2027 EP salary floor."),
    ("DISCLAIMER", "All marked EST figures are 2024-25 data or agent estimates; the search backend was throttled for several regions. This sheet informs decisions but is not a guarantee — verify each figure on the official 2027 intake page before paying fees."),
    ("COMPANION DOC", "MIM_TOP25_REPORT.md (full 13-dimension narrative per program + source URLs). STRATEGY_2027_APPLY.md (project plan + portfolio P1-P6)."),
]
for a, b in readme:
    ws4.append([a, b])
ws4.column_dimensions["A"].width = 26
ws4.column_dimensions["B"].width = 130
for c in ws4[1]:
    c.fill = PatternFill("solid", fgColor="1D1D1F")
    c.font = Font(bold=True, color="FFFFFF")
for row in ws4.iter_rows(min_row=2):
    row[0].font = Font(bold=True, size=10)
    row[0].alignment = Alignment(vertical="top")
    row[1].font = Font(size=10)
    row[1].alignment = Alignment(wrap_text=True, vertical="top")

wb.save(OUT)
print("WROTE", OUT)
print("Top25 rows:", len(TOP25), "| NearMiss rows:", len(NEARMISS))
