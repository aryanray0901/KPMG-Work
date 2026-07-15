"""Market sizing practice prompts. Each includes a guided approach and a
reasonable estimate range for self-checking, not a single 'correct' answer,
since market sizing is about defensible reasoning, not a precise figure."""

MARKET_SIZING_PROMPTS = [
    {
        "prompt": "How many piano tuners are there in Chicago?",
        "approach": "Estimate Chicago's population, then households (or specifically households "
                    "with a piano), the fraction with pianos, how often a piano needs tuning "
                    "per year, and how many tunings one tuner can do per year. Divide total annual "
                    "tunings needed by tunings-per-tuner-per-year.",
        "range_hint": "Reasonable estimates typically land between roughly 50 and 150 piano tuners, depending on assumptions.",
    },
    {
        "prompt": "How many golf balls would fit inside a standard school bus?",
        "approach": "Estimate the interior volume of a school bus (rough dimensions), the volume "
                    "of a single golf ball, and a packing efficiency factor (spheres don't pack "
                    "perfectly, roughly 60-75% of the space is typically usable).",
        "range_hint": "Reasonable estimates typically land in the hundreds of thousands, commonly cited answers range from 400,000 to 700,000.",
    },
    {
        "prompt": "What's the annual market size for umbrellas sold in New York City?",
        "approach": "Estimate NYC population and/or annual visitors, the share who buy an "
                    "umbrella in a given year (including impulse street-vendor purchases during "
                    "rain), and an average price point. Consider that many purchases are cheap, "
                    "impulse buys from street vendors rather than planned retail purchases.",
        "range_hint": "A reasonable range is roughly $15M to $40M annually, depending on assumptions about impulse purchase frequency.",
    },
    {
        "prompt": "How many gas stations are there in the United States?",
        "approach": "Estimate the number of registered vehicles in the US, average fill-ups per "
                    "vehicle per month, and typical fill-ups a single gas station can serve per "
                    "day/month based on pump count and hours of operation.",
        "range_hint": "The actual figure is roughly 145,000; reasonable structured estimates commonly land between 100,000 and 200,000.",
    },
    {
        "prompt": "How many smartphones are sold annually in your home country?",
        "approach": "Estimate total population, the smartphone-owning share, and an average "
                    "replacement cycle (how many years before someone buys a new phone). Annual "
                    "sales roughly equals owning population divided by replacement cycle years, "
                    "plus new first-time buyers.",
        "range_hint": "Reasoning quality matters more than the exact figure here since this varies widely by country; a candidate should be able to defend each assumption.",
    },
    {
        "prompt": "How many pizzas are ordered for delivery in a mid-size city (population 500,000) on a Friday night?",
        "approach": "Estimate households in the city, the share who order delivery food on a "
                    "typical Friday night, the share of those orders that are pizza specifically "
                    "versus other cuisines, and average pizzas per order.",
        "range_hint": "Reasonable estimates typically land between 5,000 and 15,000 pizzas, depending on assumptions about delivery-ordering frequency.",
    },
    {
        "prompt": "What's the total weight of all the plastic water bottles used in a year at a large university (30,000 students)?",
        "approach": "Estimate average water bottles consumed per student per day, days in an "
                    "academic year, and the weight of an average empty plastic bottle. Multiply "
                    "through and convert to a sensible unit (tons).",
        "range_hint": "Reasonable estimates typically land between 15 and 40 tons annually, depending on consumption assumptions.",
    },
    {
        "prompt": "How many commercial flights take off in the world on an average day?",
        "approach": "This is easier to anchor with a reference point if you have one (e.g. 'I "
                    "recall roughly 100,000 flights globally per day' from general knowledge); "
                    "otherwise, build up from major world regions, estimate large/medium/small "
                    "airports per region, and average daily departures per airport size.",
        "range_hint": "The actual figure is roughly 100,000-120,000 commercial flights per day; structured estimates commonly land between 60,000 and 150,000.",
    },
]
