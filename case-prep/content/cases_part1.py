"""Mock case interviews modeled after each firm's publicly known interview
style. All companies, figures, and scenarios are fictional."""

CASES = [
    {
        "slug": "altitude-air",
        "firm_style": "mckinsey",
        "title": "Altitude Air: Declining Profitability",
        "case_type": "Profitability",
        "difficulty": "Medium",
        "prompt": (
            "Your client is Altitude Air, a mid-size regional airline operating in the "
            "northeastern United States. Profits have declined 30% over the last two years "
            "despite stable revenue. The CEO has asked us to identify the cause and recommend "
            "a path back to historical profitability. Where would you start?"
        ),
        "clarifying_info": [
            "Q: Has revenue actually been flat? A: Yes, total revenue has stayed within 2% of its level two years ago.",
            "Q: Is this industry-wide or specific to Altitude Air? A: Two competitors in the region have grown profits over the same period, so it appears specific to the client.",
            "Q: Any recent major changes? A: The client added a new maintenance hangar and expanded its fleet by 15% eighteen months ago.",
        ],
        "exhibits": [
            {
                "title": "Exhibit 1: Cost Breakdown, This Year vs. Two Years Ago ($M)",
                "content": (
                    "Category            2 Yrs Ago   This Year\n"
                    "Fuel                 180         185\n"
                    "Labor                220         235\n"
                    "Maintenance           45          98\n"
                    "Airport fees          60          62\n"
                    "Other overhead        35          38\n"
                    "Total                540         618"
                ),
            },
            {
                "title": "Exhibit 2: Maintenance Cost Detail",
                "content": (
                    "The new hangar added $30M in annual fixed lease and staffing costs. "
                    "The remaining increase in maintenance cost is spread across the expanded "
                    "fleet, with per-aircraft maintenance cost roughly flat versus two years ago. "
                    "Hangar utilization is currently at 40% of capacity."
                ),
            },
        ],
        "model_approach": (
            "This is a profitability case where revenue is flat, so the framework should move "
            "quickly to the cost side. A strong candidate would segment costs by category (as in "
            "Exhibit 1), notice maintenance more than doubled, and probe why, rather than spreading "
            "attention evenly across fuel, labor, airport fees, and maintenance. Once the hangar "
            "investment surfaces as the driver, the natural follow-up is whether the hangar was "
            "sized correctly relative to the fleet expansion, which Exhibit 2's 40% utilization "
            "figure answers: the hangar is significantly underutilized relative to its cost."
        ),
        "model_synthesis": (
            "The profit decline is driven almost entirely by the new maintenance hangar, which "
            "added $30M in fixed costs but is running at only 40% utilization, well above what the "
            "current fleet size justifies. Recommend the client explore leasing excess hangar "
            "capacity to other regional carriers or third-party maintenance customers to offset the "
            "fixed cost, and revisit whether the fleet expansion that justified the hangar is still "
            "on track. Flag as a risk: if fleet growth was delayed or scaled back, the hangar "
            "investment may need to be reassessed entirely rather than optimized around."
        ),
    },
    {
        "slug": "brewhaus-expansion",
        "firm_style": "bcg",
        "title": "BrewHaus: European Market Entry",
        "case_type": "Market Entry",
        "difficulty": "Hard",
        "prompt": (
            "BrewHaus, a US specialty coffee chain with 400 locations, is considering entering "
            "Germany as its first European market. The CEO wants a recommendation within the next "
            "two weeks. What's your initial hypothesis, and how would you test it?"
        ),
        "clarifying_info": [
            "Q: Why Germany specifically? A: Leadership sees it as the largest coffee-consuming market in continental Europe and a proof point before broader European expansion.",
            "Q: What's BrewHaus's US positioning? A: Premium price point, ethically sourced beans, fast-casual format, strong loyalty app with 8M US members.",
            "Q: Any existing presence in Europe? A: None currently. No partnerships or distribution agreements in place.",
        ],
        "exhibits": [
            {
                "title": "Exhibit 1: German Coffee Market Overview",
                "content": (
                    "Market size: ~EUR 9B annual specialty coffee spend, growing 4%/year.\n"
                    "Two large domestic chains hold a combined 35% share of the specialty segment.\n"
                    "Independent cafes hold roughly 45% share; the remainder is other chains.\n"
                    "German consumers report high price sensitivity in surveys relative to US consumers, "
                    "but also high loyalty to cafes seen as authentic/local."
                ),
            },
            {
                "title": "Exhibit 2: Real Estate & Labor Cost Comparison",
                "content": (
                    "Prime retail lease rates in target German cities run 15-20% higher than "
                    "BrewHaus's typical US locations.\n"
                    "Barista labor costs (fully loaded) run roughly 25% higher than the US, "
                    "driven by stronger labor protections and benefits requirements.\n"
                    "Store buildout costs are comparable to the US once currency-adjusted."
                ),
            },
        ],
        "model_approach": (
            "A strong BCG-style answer states an early hypothesis, for example: 'My initial "
            "hypothesis is that Germany is attractive on market size but risky on unit economics "
            "given the cost structure and local competition, let's test that.' From there, structure "
            "around market attractiveness (Exhibit 1: large but fragmented, price-sensitive, strong "
            "local loyalty), then unit economics (Exhibit 2: meaningfully higher costs), then "
            "competitive response, then entry mode. The candidate should flag the tension directly: "
            "a premium-price, high-cost US model entering a price-sensitive, loyalty-driven local "
            "market with already-higher costs is a real risk, not a formality to wave past."
        ),
        "model_synthesis": (
            "Germany is a large, attractive market on size and growth, but BrewHaus's premium US "
            "model faces two compounding risks: German consumers are more price-sensitive and "
            "loyal to local/independent cafes, while BrewHaus's cost structure (real estate, labor) "
            "runs meaningfully higher than the US. Recommend piloting with a small number of stores "
            "in one city before committing to full-scale entry, and testing a locally-adapted price "
            "point rather than porting the US premium price directly. A partnership or acquisition "
            "of a well-regarded local chain may de-risk the loyalty gap faster than organic entry."
        ),
    },
    {
        "slug": "westfield-retail",
        "firm_style": "bain",
        "title": "Westfield Retail: Store Footprint Review",
        "case_type": "Operations",
        "difficulty": "Hard",
        "prompt": (
            "Westfield Retail operates 180 home goods stores nationally. The CFO wants to know "
            "which, if any, stores should be closed to improve overall profitability. Walk me "
            "through how you'd approach this."
        ),
        "clarifying_info": [
            "Q: Is this driven by a specific event? A: The board is under pressure after two "
            "consecutive quarters of margin decline company-wide.",
            "Q: Are all stores the same format/size? A: No, there are three formats: flagship "
            "(large), standard, and small-format 'express' stores.",
        ],
        "exhibits": [
            {
                "title": "Exhibit 1: Store Performance by Format",
                "content": (
                    "Format       # Stores   Avg Revenue/Store   Avg Store Margin   % of Stores Unprofitable\n"
                    "Flagship         20         $8.2M                  14%                  5%\n"
                    "Standard        120         $3.1M                   9%                 22%\n"
                    "Express          40         $1.4M                  -3%                 55%"
                ),
            },
            {
                "title": "Exhibit 2: Express Format Detail",
                "content": (
                    "Express stores were opened over the last 3 years as a lower-capex test format "
                    "targeting urban infill locations. Average lease term remaining: 2.3 years. "
                    "Unprofitable express stores are concentrated in markets where a standard-format "
                    "store already exists within 5 miles."
                ),
            },
        ],
        "model_approach": (
            "This is a Bain-style, exhibit-heavy operations case. The candidate should read Exhibit "
            "1 quickly and state the headline: express stores are the clear problem (55% "
            "unprofitable, negative average margin), not a case of broadly declining performance. "
            "The natural next question is why, which Exhibit 2 answers: many unprofitable express "
            "stores are likely cannibalizing a nearby standard store rather than capturing "
            "incremental demand. A strong candidate proposes closing overlapping express locations "
            "specifically, not the format broadly, and checks lease terms (2.3 years remaining) as a "
            "practical constraint on how fast this can happen."
        ),
        "model_synthesis": (
            "The margin decline is concentrated in the express format, and specifically in express "
            "stores that overlap with an existing standard-format location, these are likely "
            "cannibalizing sales rather than adding incremental revenue. Recommend closing "
            "overlapping express stores as leases allow (average 2.3 years remaining), while "
            "keeping express stores in markets with no nearby standard store, since those may be "
            "capturing genuinely incremental demand. Before finalizing, verify like-for-like sales "
            "impact on the nearby standard stores after a test closure or two, to confirm the "
            "cannibalization hypothesis before committing to the full closure plan."
        ),
    },
]
