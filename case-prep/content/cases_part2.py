"""Additional mock cases, continued from cases_part1.py."""

CASES_PART2 = [
    {
        "slug": "harbor-bank",
        "firm_style": "deloitte",
        "title": "Harbor Community Bank: Deposit Growth Strategy",
        "case_type": "Growth Strategy",
        "difficulty": "Medium",
        "prompt": (
            "Harbor Community Bank is a regional bank with $4B in deposits, mostly from retail "
            "checking and savings accounts. Leadership wants to grow deposits by 20% over the next "
            "three years and has asked for a growth strategy. How would you approach this?"
        ),
        "clarifying_info": [
            "Q: What's driven this deposit growth target? A: The bank wants to fund an expanded "
            "small business lending program and needs a larger deposit base to support it.",
            "Q: Any constraints on how growth is achieved? A: Leadership wants to avoid raising "
            "interest rates on deposits significantly, since that would compress margins.",
        ],
        "exhibits": [
            {
                "title": "Exhibit 1: Current Deposit Mix",
                "content": (
                    "Segment                  Share of Deposits   YoY Growth\n"
                    "Retail checking/savings         65%              2%\n"
                    "Small business accounts          20%              8%\n"
                    "High-net-worth/wealth            15%              1%"
                ),
            },
            {
                "title": "Exhibit 2: Customer Survey Insight",
                "content": (
                    "A recent survey of small business customers found 60% use a second bank for at "
                    "least one banking need, most commonly citing better digital/mobile tools as the "
                    "reason. Harbor's small business segment has grown fastest organically despite "
                    "this gap."
                ),
            },
        ],
        "model_approach": (
            "A solid growth strategy structure here separates organic growth (more from existing "
            "segments), new segment growth, and inorganic options (branch acquisition, fintech "
            "partnership). Given the constraint on rate increases, the candidate should prioritize "
            "levers that don't rely on price. Exhibit 1 shows small business is both growing "
            "fastest and smallest as a share, a natural area to lean into. Exhibit 2 then reveals a "
            "specific, addressable gap (digital tools) rather than a vague opportunity, which is "
            "exactly the kind of detail a strong candidate should probe for rather than accepting "
            "'grow all segments a bit.'"
        ),
        "model_synthesis": (
            "Recommend concentrating growth efforts on the small business segment, which is already "
            "growing fastest and has a clear, addressable gap: 60% of small business customers use "
            "a second bank primarily due to weaker digital tools. Investing in small business "
            "digital banking capabilities (mobile deposit, cash flow tools, faster loan "
            "applications) directly targets deposit consolidation from existing customers without "
            "requiring a rate increase. This should be sized against the 20% target: capturing even "
            "a portion of the 'second bank' spend from existing small business customers may cover "
            "a meaningful share of the goal before needing to pursue new-customer acquisition or "
            "M&A."
        ),
    },
    {
        "slug": "meridian-health",
        "firm_style": "pwc",
        "title": "Meridian Health System: Operating Cost Reduction",
        "case_type": "Operations",
        "difficulty": "Medium",
        "prompt": (
            "Meridian Health System operates five hospitals and is facing margin pressure from "
            "rising labor costs and flat reimbursement rates. The CFO has asked for a plan to "
            "reduce operating costs by 8% without affecting patient care quality. Where would you "
            "start?"
        ),
        "clarifying_info": [
            "Q: What's the current cost breakdown roughly? A: Labor is about 55% of operating "
            "costs, supplies/pharmaceuticals about 25%, facilities and overhead the remainder.",
            "Q: Any specific quality metrics we need to protect? A: Patient readmission rates and "
            "patient satisfaction scores are both tied to reimbursement and must not decline.",
        ],
        "exhibits": [
            {
                "title": "Exhibit 1: Labor Cost Detail",
                "content": (
                    "Nursing overtime hours have increased 35% over two years, driven largely by "
                    "understaffing on night shifts at three of the five hospitals. Agency (temporary) "
                    "nursing costs, used to fill these gaps, run roughly 60% more per hour than "
                    "staff nursing."
                ),
            },
            {
                "title": "Exhibit 2: Supply Chain Detail",
                "content": (
                    "Each of the five hospitals currently negotiates medical supply contracts "
                    "independently rather than through a system-wide group purchasing arrangement. "
                    "Peer health systems of similar size report 10-15% lower per-unit supply costs "
                    "using centralized purchasing."
                ),
            },
        ],
        "model_approach": (
            "With labor at 55% of costs, that's the natural place to focus first, but the candidate "
            "should avoid the generic 'reduce headcount' instinct given the quality constraint. "
            "Exhibit 1 points to a specific, addressable root cause: understaffing is driving "
            "expensive overtime and even more expensive agency nursing, a staffing and scheduling "
            "problem, not simply a 'cut costs' problem, and one that plausibly improves quality "
            "(fewer overworked night shifts) rather than harming it. Exhibit 2 offers a second, "
            "independent lever (procurement) that doesn't touch patient care at all, worth surfacing "
            "as a complementary rather than competing recommendation."
        ),
        "model_synthesis": (
            "Two complementary levers: first, address the nursing staffing gap directly by hiring "
            "additional staff nurses for understaffed night shifts at the three affected hospitals, "
            "the increased base cost is likely more than offset by eliminating the 60% premium on "
            "agency nursing and reducing overtime, while also plausibly improving care quality "
            "rather than risking it. Second, consolidate medical supply purchasing across all five "
            "hospitals into a single group arrangement, which peer benchmarks suggest could reduce "
            "supply costs 10-15% with no impact on care delivery. Recommend sizing both levers "
            "against the 8% target before determining whether further measures are needed."
        ),
    },
    {
        "slug": "ironclad-manufacturing",
        "firm_style": "ey",
        "title": "Ironclad Manufacturing: Acquisition Evaluation",
        "case_type": "M&A",
        "difficulty": "Hard",
        "prompt": (
            "Ironclad Manufacturing, an industrial parts maker, is considering acquiring a smaller "
            "competitor, Bolt & Frame, for $200M. The CEO wants to know if this is a good deal. "
            "How would you evaluate it?"
        ),
        "clarifying_info": [
            "Q: What's the strategic rationale being considered? A: Bolt & Frame has a strong "
            "position in the automotive fasteners niche that Ironclad doesn't currently serve.",
            "Q: Is $200M roughly in line with Bolt & Frame's standalone value? A: Yes, that's "
            "roughly 8x their current EBITDA of $25M, in line with recent comparable deals.",
        ],
        "exhibits": [
            {
                "title": "Exhibit 1: Synergy Estimate",
                "content": (
                    "Cost synergies (shared distribution, overhead consolidation): $8M/year run-rate, "
                    "expected to take 18 months to fully realize.\n"
                    "Revenue synergies (cross-selling Ironclad's broader catalog to Bolt & Frame's "
                    "automotive customers): estimated $5M/year, but management has low confidence "
                    "in this estimate."
                ),
            },
            {
                "title": "Exhibit 2: Integration Risk Factors",
                "content": (
                    "Bolt & Frame's ERP and manufacturing systems are on a different platform than "
                    "Ironclad's. A prior Ironclad acquisition (three years ago) saw system "
                    "integration take twice as long as planned, contributing to customer service "
                    "issues during the transition."
                ),
            },
        ],
        "model_approach": (
            "A strong candidate values the deal in two layers: standalone value (given as roughly "
            "in line with the purchase price) plus synergies, then discounts for realistic execution "
            "risk rather than taking Exhibit 1's numbers at face value. The prompt explicitly flags "
            "low management confidence in the revenue synergy estimate, a strong candidate should "
            "weight that appropriately rather than including it at full value. Exhibit 2 is a direct "
            "signal to discount the cost synergy timeline too: this isn't Ironclad's first "
            "integration, and their own history suggests these timelines tend to slip."
        ),
        "model_synthesis": (
            "At $200M against $25M EBITDA (8x), the deal is roughly fairly valued on a standalone "
            "basis, so the case for approval rests on synergies. The $8M cost synergy has "
            "reasonable support but a track record of slipping (per Ironclad's own prior "
            "acquisition), and the $5M revenue synergy carries explicitly low management "
            "confidence and shouldn't be counted on to justify the price. Recommend proceeding, but "
            "at a valuation that reflects only a risk-adjusted portion of synergies, and pushing for "
            "a longer integration runway and a dedicated systems integration plan given the prior "
            "acquisition's history, rather than assuming this one will go more smoothly."
        ),
    },
    {
        "slug": "cloudsuite-pricing",
        "firm_style": "kpmg",
        "title": "CloudSuite: SaaS Pricing Decision",
        "case_type": "Pricing",
        "difficulty": "Medium",
        "prompt": (
            "CloudSuite sells project management software to mid-size businesses on a per-seat "
            "subscription model, currently $40/user/month. The VP of Sales wants to raise the "
            "price to $50/user/month to boost revenue. Should they?"
        ),
        "clarifying_info": [
            "Q: How long has the price been $40? A: Three years, with no increase despite adding "
            "significant new features over that time.",
            "Q: What does the competitive landscape look like? A: Two direct competitors price "
            "similarly ($38-42/user/month); one premium competitor prices at $65/user/month with "
            "a more limited feature set.",
        ],
        "exhibits": [
            {
                "title": "Exhibit 1: Customer Segment Detail",
                "content": (
                    "Segment            % of Customers   Price Sensitivity (survey)\n"
                    "Small teams (<20)        45%              High\n"
                    "Mid-size (20-100)        40%              Medium\n"
                    "Larger accounts (100+)   15%              Low"
                ),
            },
            {
                "title": "Exhibit 2: Churn Risk Analysis",
                "content": (
                    "Internal analysis estimates a $10 price increase would cause 12-18% churn "
                    "among small teams, 5-8% among mid-size, and under 2% among larger accounts, "
                    "based on a recent price test in a similar customer base."
                ),
            },
        ],
        "model_approach": (
            "A strong pricing case doesn't answer 'raise or don't raise' as a single yes/no, it "
            "recognizes the customer base isn't homogeneous. Exhibit 1 sets up the segmentation, "
            "and Exhibit 2 quantifies that price sensitivity varies a lot by segment. The candidate "
            "should propose modeling net revenue impact by segment (accounting for churn) rather "
            "than assuming a flat price increase is uniformly good or bad, and should notice that a "
            "single across-the-board price change is probably not optimal when segments respond so "
            "differently."
        ),
        "model_synthesis": (
            "A flat price increase to $50 across all customers risks meaningful churn in the "
            "small-team segment (45% of customers, 12-18% estimated churn) for a segment that's "
            "also most price-sensitive, while larger accounts could likely absorb the increase with "
            "minimal churn. Recommend a segmented approach: raise price for mid-size and larger "
            "accounts where churn risk is low and the increase is likely net-positive for revenue, "
            "while holding or offering a lower-tier price point for small teams to protect that "
            "segment, or introducing a scaled-down feature tier at the current price point to "
            "preserve the entry-level customer base. This also fits the competitive landscape, "
            "where CloudSuite has room to move toward the premium competitor's price point for "
            "higher-value segments without abandoning price-sensitive smaller customers entirely."
        ),
    },
    {
        "slug": "ridewell-sizing",
        "firm_style": "mckinsey",
        "title": "RideWell: New City Market Sizing",
        "case_type": "Market Sizing",
        "difficulty": "Medium",
        "prompt": (
            "RideWell, a ride-sharing company, is deciding whether to launch in a new mid-size "
            "city of 1.2 million people. Before going further, they want an estimate of the "
            "annual market size for ride-sharing in that city. How would you size it?"
        ),
        "clarifying_info": [
            "Q: Should I assume this city is similar to cities RideWell already operates in? "
            "A: Yes, assume broadly similar demographics and transit infrastructure to RideWell's "
            "existing mid-size markets unless a specific reason suggests otherwise.",
        ],
        "exhibits": [],
        "model_approach": (
            "A strong market sizing answer states a clear structure before calculating: total "
            "population, narrow to the addressable population (adults, likely excluding those "
            "without smartphones or reliable payment methods), estimate the share who'd use "
            "ride-sharing at all, then estimate rides per user per month and average fare, and "
            "annualize. The candidate should state each assumption explicitly and round to clean "
            "numbers rather than false precision, for example '1.2M people, call it 900K adults, "
            "of whom maybe 40% are realistic ride-share users given smartphone/payment access and "
            "urban density' rather than an unexplained final number."
        ),
        "model_synthesis": (
            "Example structure and output: 1.2M population -> ~900K adults -> ~40% realistic "
            "ride-share users (360K) -> ~3 rides/month average among users -> ~13M rides/year -> "
            "at an average fare of $12, roughly $155M in annual gross bookings. The specific number "
            "matters less than a clean, defensible structure and reasonable, clearly-stated "
            "assumptions at each step; a strong candidate would also sanity-check the result "
            "against a known reference point if one is available (e.g. a similar city RideWell "
            "already operates in) and adjust if the estimate seems far off."
        ),
    },
    {
        "slug": "civicharge",
        "firm_style": "bcg",
        "title": "CivicCharge: Public EV Infrastructure Investment",
        "case_type": "Public Sector / Growth",
        "difficulty": "Hard",
        "prompt": (
            "A mid-size city government is deciding whether to invest $40M in public electric "
            "vehicle charging infrastructure over the next five years. The transportation "
            "director has asked for a recommendation. How would you think about this?"
        ),
        "clarifying_info": [
            "Q: What's the goal, revenue, emissions reduction, or something else? A: Primarily "
            "emissions reduction and supporting EV adoption; the city isn't expecting charging "
            "fees to cover the full cost.",
            "Q: Is private investment in charging infrastructure happening already? A: Some, "
            "concentrated in higher-income neighborhoods and shopping centers.",
        ],
        "exhibits": [
            {
                "title": "Exhibit 1: Current EV Adoption & Charging Access",
                "content": (
                    "Citywide EV adoption: 6% of registered vehicles, growing ~1.5 points/year.\n"
                    "Public charging access is uneven: neighborhoods with existing private charging "
                    "investment have 3x the charger density of lower-income neighborhoods."
                ),
            },
            {
                "title": "Exhibit 2: Resident Survey",
                "content": (
                    "Among residents without home charging access (mostly renters and those in "
                    "multi-unit housing), 65% cite lack of convenient public charging as a top "
                    "barrier to considering an EV."
                ),
            },
        ],
        "model_approach": (
            "Given the stated goal is emissions/adoption rather than revenue, the candidate should "
            "avoid defaulting to a pure ROI/payback framework and instead structure around: where "
            "is the barrier to adoption concentrated, is public investment additive to what's "
            "already happening privately, and what's the most cost-effective way to close the gap. "
            "Exhibit 1 shows uneven access concentrated by income/neighborhood, and Exhibit 2 gives "
            "a direct causal signal (lack of charging is a top-cited adoption barrier for exactly "
            "the underserved group). A strong candidate connects these explicitly rather than "
            "treating them as separate facts."
        ),
        "model_synthesis": (
            "Private investment is already addressing higher-income neighborhoods, so public "
            "investment is most additive, and most aligned with the equity and adoption goals, if "
            "concentrated in underserved neighborhoods with low existing charger density, where "
            "survey data shows charging access is a specific, addressable barrier to EV adoption "
            "for residents without home charging (largely renters and multi-unit housing residents). "
            "Recommend prioritizing the $40M toward these areas rather than spreading it evenly "
            "citywide or over-indexing on high-visibility/high-traffic locations that private "
            "investment is already covering. Recommend a phased rollout with adoption tracking "
            "in early neighborhoods to validate the approach before committing the full budget."
        ),
    },
]
