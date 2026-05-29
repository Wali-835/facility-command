import { EmailTemplate } from '../types';

export const EMAIL_TEMPLATES: EmailTemplate[] = [
  {
    id: 'intro-01',
    name: 'Product Introduction',
    category: 'introduction',
    subject: 'Transform Your Facility Cleaning Operations with Fiorentini',
    body: `Dear {{firstName}},

Managing a large facility comes with its fair share of challenges — and keeping floors clean efficiently shouldn't be one of them.

We represent Fiorentini, Italy's premier industrial cleaning machine manufacturer, and we bring world-class floor cleaning solutions to warehouses and facilities like yours.

Why Fiorentini?

✅ Clean up to 6,000 m²/hour — no more slow, labour-intensive cleaning
✅ 50% lower operating costs compared to manual cleaning
✅ Built for industrial environments — warehouses, logistics centres, manufacturing plants
✅ Intuitive controls — minimal operator training required
✅ Energy-efficient lithium battery models available

Whether you manage a 5,000 m² distribution centre or a 50,000 m² manufacturing facility, Fiorentini has a solution that fits.

Our current lineup includes:
• Walk-behind scrubber-dryers for tight spaces and smaller areas
• Ride-on scrubber-dryers for large open floors
• Industrial sweepers for heavy debris and outdoor areas
• Combined sweeper-scrubbers for maximum versatility

I'd love to schedule a quick 15-minute call to understand your specific needs and show you which Fiorentini model would deliver the best ROI for your facility.

Best regards,
{{senderName}}`,
  },
  {
    id: 'roi-01',
    name: 'ROI & Cost Savings',
    category: 'roi',
    subject: 'How Facility Managers Cut Cleaning Costs by 40% with Fiorentini',
    body: `Dear {{firstName}},

A logistics company with three warehouses was spending €180,000 per year on manual cleaning labour. After switching to Fiorentini ride-on scrubbers, their annual cost dropped to €85,000 — a 53% saving in year one.

The secret? Industrial-grade cleaning that does in 2 hours what used to take 8.

Here's what makes Fiorentini stand out for warehouse and facility managers:

PRODUCTIVITY
• Cover up to 6,000 m²/hour with ride-on models
• Consistent cleaning quality — no "tired at end of shift" variance
• One operator replaces 4–6 manual cleaners

TOTAL COST OF OWNERSHIP
• Lithium battery models: no fuel costs, lower maintenance
• Parts and service available within 24 hours
• 3-year warranty on key components

COMPLIANCE & SAFETY
• HACCP-compatible models for food-grade facilities
• ATEX-certified options for hazardous environments
• Floors dry in under 5 minutes — significantly reduces slip hazards

Would you like to see the numbers for your specific facility? I can put together a custom ROI calculation based on your floor area and current cleaning spend.

Regards,
{{senderName}}`,
  },
  {
    id: 'demo-01',
    name: 'Free On-Site Demo',
    category: 'demo',
    subject: 'Free Fiorentini Demo at Your Facility — No Obligation',
    body: `Dear {{firstName}},

Actions speak louder than brochures.

That's why we're offering facility managers in your area a completely free, no-obligation on-site demonstration of our Fiorentini industrial cleaning machines.

Here's what you'll see:

🎯 Live cleaning performance on your actual floors
📊 Real-time speed and coverage metrics
💡 Side-by-side comparison with your current cleaning method
💰 On-the-spot ROI calculation for your facility

No strings attached. If Fiorentini isn't the right fit for your operation, we'll tell you so.

Facility managers who've attended consistently say it's the most useful 45 minutes they've spent all quarter.

To reserve your demo slot, simply reply to this email with your preferred date and time, and I'll confirm within the hour.

Best,
{{senderName}}`,
  },
  {
    id: 'followup-01',
    name: 'Follow-Up / Re-engagement',
    category: 'followup',
    subject: 'Still thinking about upgrading your facility cleaning? A quick update',
    body: `Dear {{firstName}},

I wanted to follow up and share some new information that might be relevant to your facility.

What's new from Fiorentini:

We've recently expanded our lineup with models designed specifically for the demands of modern logistics and manufacturing environments — including narrower ride-on configurations for facilities with racking at 2.8m aisle widths, and new lithium battery platforms that offer up to 8 hours of continuous runtime.

We're also currently running a promotion that includes:
✅ Free installation and operator training (valued at €2,500)
✅ 12-month preventive maintenance package included
✅ Flexible payment terms available

I know your time is valuable, so I'll keep this brief: if upgrading your cleaning operations is on your radar for this year, now is a great time to have a conversation.

If this isn't a priority right now, no problem at all — just let me know and I'll follow up at a more appropriate time.

Kind regards,
{{senderName}}`,
  },
];
