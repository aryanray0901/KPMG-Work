"""Reference cards for standard case interview frameworks."""

FRAMEWORKS = [
    {
        "slug": "profitability",
        "name": "Profitability Framework",
        "when": "Profit is declining, flat, or below target, and the case doesn't specify why.",
        "structure": [
            "Profit = Revenue - Costs. Start by asking which side (or both) has changed.",
            "Revenue = Price x Volume. Break volume into # of customers x purchase frequency x units per purchase, or by segment/region/product line.",
            "Costs = Fixed + Variable. Break down by major cost category (COGS, labor, overhead, marketing) and check for step-changes (e.g. a new facility, new regulation).",
            "Look for a specific driver before recommending: a single segment, region, product, or time period where the change concentrated.",
        ],
        "pitfall": "Listing every possible revenue and cost driver without prioritizing. State which branch you think is more likely to matter and why, before diving in.",
    },
    {
        "slug": "market-entry",
        "name": "Market Entry Framework",
        "when": "A client is deciding whether to enter a new market, geography, or product category.",
        "structure": [
            "Market attractiveness: size, growth rate, profitability, regulatory environment.",
            "Competitive landscape: who's already there, market shares, how they'd likely respond to entry.",
            "Client capabilities and fit: does the client have the assets, brand, or capabilities to compete, or would they need to build/buy/partner?",
            "Entry mode and financial case: organic build, acquisition, joint venture, or partnership, and the rough economics of each.",
        ],
        "pitfall": "Spending most of the case on market sizing and running out of time on the 'should they enter, and how' question, which is usually the actual client question.",
    },
    {
        "slug": "ma-synergies",
        "name": "M&A / Synergies Framework",
        "when": "A client is considering acquiring, merging with, or being acquired by another company.",
        "structure": [
            "Strategic rationale: why this deal, market access, capability gap, scale, defensive move.",
            "Standalone valuation of the target, then synergies: revenue synergies (cross-sell, pricing power) and cost synergies (overhead, procurement, facility consolidation).",
            "Integration risk: cultural fit, systems integration, key talent retention, customer disruption during integration.",
            "Deal structure and price: is the synergy-adjusted value above the likely purchase price, including a reasonable margin of safety.",
        ],
        "pitfall": "Treating synergies as free upside without discounting for integration risk and realistic execution timelines.",
    },
    {
        "slug": "growth-strategy",
        "name": "Growth Strategy Framework",
        "when": "A client wants to grow revenue or market share and the path isn't specified.",
        "structure": [
            "Organic growth: sell more to existing customers, win new customers in existing markets, or launch new products.",
            "New market growth: geographic expansion or new customer segments.",
            "Inorganic growth: acquisition or partnership to acquire capability, customers, or market access faster than organic growth would allow.",
            "Prioritize by a rough size-of-prize and feasibility for each option before recommending.",
        ],
        "pitfall": "Proposing every growth lever with equal weight instead of sizing and prioritizing which one or two are worth pursuing first.",
    },
    {
        "slug": "pricing",
        "name": "Pricing Framework",
        "when": "The case centers on whether/how to change price for a product or service.",
        "structure": [
            "Cost-based floor: what's the minimum viable price given cost structure and margin targets.",
            "Value-based ceiling: what's the customer's willingness to pay, based on the value delivered relative to alternatives.",
            "Competitive positioning: where do comparable offerings sit, and does the client want to lead or follow on price.",
            "Elasticity and volume impact: model how a price change likely affects volume, and whether the net revenue/margin effect is positive.",
        ],
        "pitfall": "Recommending a price change without addressing how volume will likely respond, price and volume move together, not independently.",
    },
    {
        "slug": "operations",
        "name": "Operations / Cost Reduction Framework",
        "when": "A client needs to reduce costs or improve operational efficiency without a specified area.",
        "structure": [
            "Map the value chain or process end to end to see where cost or time is concentrated.",
            "Segment costs into direct/COGS, labor, overhead, and evaluate each against benchmarks or best practice.",
            "Identify quick wins (procurement renegotiation, waste reduction) versus structural changes (automation, footprint consolidation, org redesign).",
            "Weigh cost-reduction impact against implementation risk, cost, and time, and against any effect on quality or customer experience.",
        ],
        "pitfall": "Recommending aggressive cost cuts without flagging the risk to quality, morale, or customer experience.",
    },
    {
        "slug": "three-cs",
        "name": "3Cs (Company, Customers, Competitors)",
        "when": "A broad, open-ended strategic question where a narrower framework doesn't obviously fit.",
        "structure": [
            "Company: internal capabilities, financial position, strategic priorities, and constraints.",
            "Customers: who they are, what they value, how that's shifting.",
            "Competitors: who else is in the space, their relative strengths, and likely response to whatever the client does.",
        ],
        "pitfall": "Using 3Cs as a generic catch-all without tailoring each bucket's questions to the specific case question.",
    },
    {
        "slug": "porters-five-forces",
        "name": "Porter's Five Forces",
        "when": "Assessing overall industry attractiveness or a competitive dynamics question.",
        "structure": [
            "Threat of new entrants: how easy is it for a new competitor to enter this market.",
            "Bargaining power of suppliers: can suppliers dictate terms, or are there many substitutable suppliers.",
            "Bargaining power of buyers: can customers negotiate down price/terms, or is demand fragmented.",
            "Threat of substitutes: are there alternative products or services that solve the same customer need.",
            "Competitive rivalry: how intense is competition among existing players.",
        ],
        "pitfall": "Running through all five forces exhaustively when the case question is really about just one or two of them.",
    },
]


def get_framework(slug):
    for f in FRAMEWORKS:
        if f["slug"] == slug:
            return f
    return None
