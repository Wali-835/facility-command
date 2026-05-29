export const MACHINES = [
  {
    id: 'walk-behind-scrubber',
    label: 'Walk-Behind Scrubber-Dryer',
    coverage: '1,200–2,500 m²/hr',
    bestFor: 'smaller facilities, tight spaces, aisles',
    keyBenefits: [
      'compact design fits narrow aisles',
      'floor dry in under 5 minutes',
      '60% reduction in water and chemical usage',
      'one operator replaces 3–4 manual cleaners',
      'intuitive controls — trained in under 30 minutes',
    ],
    painSolved: 'manual mopping in congested warehouse aisles',
  },
  {
    id: 'ride-on-scrubber',
    label: 'Ride-On Scrubber-Dryer',
    coverage: '4,000–6,000 m²/hr',
    bestFor: 'large warehouses, distribution centres, logistics hubs',
    keyBenefits: [
      'covers 6,000 m²/hr — a full shift in 2 hours',
      'operator comfort for 8-hour shifts',
      'lithium battery option for zero emissions indoors',
      'consistent results regardless of operator',
      'reduces cleaning staff headcount by up to 70%',
    ],
    painSolved: 'slow, expensive cleaning across vast floor areas',
  },
  {
    id: 'industrial-sweeper',
    label: 'Industrial Sweeper',
    coverage: '3,000–5,000 m²/hr',
    bestFor: 'manufacturing plants, outdoor yards, heavy debris environments',
    keyBenefits: [
      'handles large debris, dust, and granular waste',
      'dust-suppression system keeps air quality safe',
      'suitable for indoor and outdoor use',
      'reduces forklift contamination from floor debris',
      'lowers dust-related maintenance on equipment',
    ],
    painSolved: 'debris accumulation in manufacturing and production floors',
  },
  {
    id: 'combined-sweeper-scrubber',
    label: 'Combined Sweeper-Scrubber',
    coverage: '3,500–5,500 m²/hr',
    bestFor: 'mixed environments, food production, multi-use facilities',
    keyBenefits: [
      'sweeps and scrubs in a single pass — half the time',
      'ideal for facilities with both dust and liquid soiling',
      'HACCP-compatible models for food-grade environments',
      'reduces number of machines needed by 50%',
      'lower total cost of ownership',
    ],
    painSolved: 'managing two separate cleaning processes across large areas',
  },
  {
    id: 'lithium-battery-series',
    label: 'Lithium Battery Series',
    coverage: 'Up to 6,000 m²/hr',
    bestFor: 'facilities with indoor air quality requirements, sustainability targets',
    keyBenefits: [
      'zero emissions — safe for enclosed environments',
      'up to 8 hours runtime on a single charge',
      '30% lower energy cost vs. traditional models',
      'rapid opportunity charging — no downtime between shifts',
      'aligns with ESG and sustainability reporting goals',
    ],
    painSolved: 'indoor air quality concerns and rising energy costs',
  },
];

export const POST_TYPES = [
  { id: 'product-spotlight', label: 'Product Spotlight' },
  { id: 'industry-insight', label: 'Industry Insight' },
  { id: 'client-success', label: 'Client Success Story' },
  { id: 'maintenance-tip', label: 'Maintenance / Efficiency Tip' },
  { id: 'problem-solution', label: 'Problem → Solution' },
];

export const TONES = [
  { id: 'professional', label: 'Professional' },
  { id: 'conversational', label: 'Conversational' },
  { id: 'data-driven', label: 'Data-Driven' },
];

export const HASHTAGS: Record<string, string[]> = {
  general: ['#FacilityManagement', '#WarehouseOperations', '#IndustrialCleaning', '#Fiorentini'],
  'walk-behind-scrubber': ['#FloorCare', '#CleaningEquipment', '#SmallWarehouse'],
  'ride-on-scrubber': ['#WarehouseEfficiency', '#LogisticsCentre', '#FleetManagement'],
  'industrial-sweeper': ['#Manufacturing', '#IndustrialSafety', '#DustControl'],
  'combined-sweeper-scrubber': ['#FoodSafety', '#HACCP', '#OperationalExcellence'],
  'lithium-battery-series': ['#Sustainability', '#ESG', '#GreenOperations', '#ZeroEmissions'],
};

type ToneKey = 'professional' | 'conversational' | 'data-driven';

interface PostGeneratorParams {
  postType: string;
  machineId: string;
  tone: string;
  customPoint?: string;
}

export function generateLinkedInPost(params: PostGeneratorParams): string {
  const { postType, machineId, tone, customPoint } = params;
  const machine = MACHINES.find(m => m.id === machineId) ?? MACHINES[0]!;
  const t = (tone as ToneKey) ?? 'professional';

  const benefit1 = machine.keyBenefits[0] ?? '';
  const benefit2 = machine.keyBenefits[1] ?? '';
  const benefit3 = machine.keyBenefits[2] ?? '';
  const benefit4 = machine.keyBenefits[3] ?? '';

  const hashtags = [
    ...HASHTAGS.general,
    ...(HASHTAGS[machineId] ?? []),
  ].join(' ');

  const customLine = customPoint
    ? `\nOne thing worth highlighting: ${customPoint}\n`
    : '';

  if (postType === 'product-spotlight') {
    const hooks: Record<ToneKey, string> = {
      professional: `Most facility managers I speak with are still relying on manual methods for daily floor maintenance.`,
      conversational: `Quick question for facility and warehouse managers in my network:`,
      'data-driven': `Facilities using manual floor cleaning spend an average of 4–6 staff hours per cycle — with inconsistent results.`,
    };
    return `${hooks[t]}

The result is predictable:
→ High labour cost that grows every year
→ Inconsistent cleaning quality
→ Wet floors that create slip hazards
→ Staff fatigue leading to shortcuts

Here's what changes with the Fiorentini ${machine.label}:

✅ ${benefit1.charAt(0).toUpperCase() + benefit1.slice(1)}
✅ ${benefit2.charAt(0).toUpperCase() + benefit2.slice(1)}
✅ ${benefit3.charAt(0).toUpperCase() + benefit3.slice(1)}
✅ ${benefit4.charAt(0).toUpperCase() + benefit4.slice(1)}

Coverage: ${machine.coverage} — best suited for ${machine.bestFor}.
${customLine}
If floor cleanliness and operational efficiency are priorities for your facility, I'd be happy to walk you through what Fiorentini can do for your specific operation.

Drop a comment or send me a message.

${hashtags}`;
  }

  if (postType === 'industry-insight') {
    const hooks: Record<ToneKey, string> = {
      professional: `The industrial cleaning sector is undergoing a significant shift — and facility managers who adapt early are gaining a measurable competitive advantage.`,
      conversational: `Something I keep seeing across warehouses and distribution centres:`,
      'data-driven': `Industry data shows that facilities investing in automated cleaning equipment reduce total cleaning costs by 35–55% within the first 24 months.`,
    };
    return `${hooks[t]}

Three trends reshaping facility management in 2025:

1. Labour cost pressure — cleaning staff costs have risen 20–30% in the past three years. Automation is no longer a luxury; it's a necessity.

2. Safety and compliance — slip-and-fall incidents from wet floors account for a significant share of workplace injuries. Equipment that dries floors in under 5 minutes is increasingly a compliance requirement.

3. Sustainability targets — ESG reporting is driving demand for battery-powered, low-chemical cleaning solutions.

The Fiorentini ${machine.label} addresses all three:
• ${benefit1.charAt(0).toUpperCase() + benefit1.slice(1)}
• ${benefit2.charAt(0).toUpperCase() + benefit2.slice(1)}
• ${benefit3.charAt(0).toUpperCase() + benefit3.slice(1)}
${customLine}
Facilities that have made the switch consistently report ROI within 12–18 months.

Curious what this looks like for your operation? Let's connect.

${hashtags} #OperationsManagement #FutureOfWork`;
  }

  if (postType === 'client-success') {
    const hooks: Record<ToneKey, string> = {
      professional: `A distribution centre managing over 20,000 m² across two shifts was struggling with cleaning costs and inconsistent floor hygiene standards.`,
      conversational: `Here's a real story from a facility manager we worked with recently.`,
      'data-driven': `Before: €160,000/year in cleaning labour. After switching to Fiorentini: €72,000/year. Same facility. Same square footage. 55% cost reduction.`,
    };
    return `${hooks[t]}

The challenge:
• Manual cleaning required 6 staff per shift
• Floor hygiene was inconsistent — a recurring issue in audits
• Wet floors were causing near-miss slip incidents
• Cleaning was consuming 3–4 hours of each shift

The solution: Fiorentini ${machine.label}

Why they chose it:
→ ${benefit1.charAt(0).toUpperCase() + benefit1.slice(1)}
→ ${benefit2.charAt(0).toUpperCase() + benefit2.slice(1)}
→ ${benefit3.charAt(0).toUpperCase() + benefit3.slice(1)}
${customLine}
The outcome after 6 months:
✅ Cleaning team reduced from 6 to 2 operators
✅ Zero slip incidents related to wet floors
✅ Cleaning cycle time cut from 3.5 hours to 90 minutes
✅ Audit scores improved significantly

The facility manager told me: "I wish we'd done this three years ago."

If you're managing a similar challenge, I'm happy to share how we approached it.

${hashtags} #CaseStudy #ROI`;
  }

  if (postType === 'maintenance-tip') {
    const hooks: Record<ToneKey, string> = {
      professional: `5 signs your facility's floor cleaning operation needs an upgrade — and what to do about each one.`,
      conversational: `Sharing something practical for facility managers today:`,
      'data-driven': `Analysis across 50+ warehouse facilities shows these 5 indicators reliably predict rising cleaning costs — and avoidable downtime.`,
    };
    return `${hooks[t]}

1. Your cleaning cycle takes more than 3 hours
Modern ride-on and walk-behind scrubbers cover the same floor area in under 90 minutes. If you're over 3 hours, you're leaving productivity on the table.

2. You have more than 4 staff dedicated to daily floor cleaning
For most facilities under 10,000 m², 1–2 operators with the right equipment is sufficient. More than that means labour is doing what machinery should.

3. You've had slip-and-fall near-misses on wet floors
Proper scrubber-dryers leave floors dry in under 5 minutes. Extended drying times are an equipment problem, not an operational one.

4. Your cleaning results vary by shift or by operator
Equipment delivers consistent results. If your floors look different depending on who cleaned them, the process isn't reliable.

5. You haven't benchmarked your cleaning cost per m² recently
The Fiorentini ${machine.label} brings cost per m² down to a fraction of manual cleaning — ${machine.coverage} coverage per hour.
${customLine}
If any of these resonate, happy to share what a modern facility cleaning setup looks like.

${hashtags} #FacilityTips #OperationsManagement`;
  }

  if (postType === 'problem-solution') {
    const hooks: Record<ToneKey, string> = {
      professional: `The biggest hidden cost in warehouse operations isn't rent, energy, or even labour — it's inefficient floor maintenance.`,
      conversational: `Something I hear from facility managers almost every week:`,
      'data-driven': `For a 15,000 m² facility, the difference between manual cleaning and an automated Fiorentini solution is typically €60,000–€90,000 per year.`,
    };
    return `${hooks[t]}

The problem:
${machine.painSolved.charAt(0).toUpperCase() + machine.painSolved.slice(1)} is costing facilities more than most managers realise — not just in labour, but in compliance risk, equipment wear from debris, and staff turnover from physically demanding work.

The old approach:
❌ Manual mops and buckets — slow, inconsistent, leave floors wet for 30–60 minutes
❌ Hiring more cleaning staff — costs scale linearly with floor area
❌ Outsourcing to cleaning contractors — high cost, low accountability
❌ Accepting "good enough" — until an audit or incident forces a change

The Fiorentini approach:

The ${machine.label} was built for exactly this environment.

→ ${benefit1.charAt(0).toUpperCase() + benefit1.slice(1)}
→ ${benefit2.charAt(0).toUpperCase() + benefit2.slice(1)}
→ ${benefit3.charAt(0).toUpperCase() + benefit3.slice(1)}
→ ${benefit4.charAt(0).toUpperCase() + benefit4.slice(1)}

Coverage: ${machine.coverage} | Best for: ${machine.bestFor}
${customLine}
The ROI is typically clear within 12–18 months. Some facilities see it in under 6.

If this matches a challenge you're facing, let's have a conversation.

${hashtags} #ProblemSolving #WarehouseManagement`;
  }

  return '';
}
