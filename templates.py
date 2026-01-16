"""
templates.py
Template registry and management for Analog Garage Workbench.
Supports multiple templates with easy addition and iteration.
"""

from typing import Dict, List, Optional

# =============================================================================
# EXPLICIT EXPORTS
# =============================================================================

__all__ = [
    'PromptTemplate',
    'TemplateRegistry', 
    'template_registry',
    'get_template',
    'list_templates',
    'render_template',
    'build_enhanced_prompt',
]

# =============================================================================
# GEOGRAPHIC BREAKDOWN SECTIONS (Used by V3 template)
# =============================================================================

GEOGRAPHIC_BREAKDOWN_GLOBAL = """
**GEOGRAPHIC BREAKDOWN (Global Scope):**

Provide detailed breakdown of TAM and value across major regions:

| Region | Population/Units | % of Global TAM | Regional Factors | Market Readiness |
|--------|-----------------|-----------------|------------------|------------------|
| United States | [number] | [%] | [Regulatory, reimbursement, adoption factors] | [High/Med/Low] |
| EMEA (Europe, Middle East, Africa) | [number] | [%] | [Key market characteristics] | [High/Med/Low] |
| Japan | [number] | [%] | [Regulatory, cultural factors] | [High/Med/Low] |
| China | [number] | [%] | [Market access, regulatory factors] | [High/Med/Low] |
| Rest of World | [number] | [%] | [Emerging market considerations] | [High/Med/Low] |
| **GLOBAL TOTAL** | **[number]** | **100%** | | |

**Regional Value Multipliers:**
(Account for pricing differences, healthcare spending, purchasing power)

| Region | Base Value/Unit | Regional Multiplier | Adjusted Value/Unit |
|--------|-----------------|--------------------|--------------------|
| United States | $[amount] | 1.0 (baseline) | $[amount] |
| EMEA | $[amount] | [0.X] | $[amount] |
| Japan | $[amount] | [0.X] | $[amount] |
| China | $[amount] | [0.X] | $[amount] |
| Rest of World | $[amount] | [0.X] | $[amount] |
"""

GEOGRAPHIC_VALUE_GLOBAL = """
**6C. Value by Geographic Region (100% Penetration):**

| Region | Regional TAM | Value/Unit | Total Regional Value | % of Global Value | Monetizable Value |
|--------|--------------|------------|---------------------|-------------------|-------------------|
| United States | [units] | $[amount] | $[amount] | [%] | $[amount] |
| EMEA | [units] | $[amount] | $[amount] | [%] | $[amount] |
| Japan | [units] | $[amount] | $[amount] | [%] | $[amount] |
| China | [units] | $[amount] | $[amount] | [%] | $[amount] |
| Rest of World | [units] | $[amount] | $[amount] | [%] | $[amount] |
| **GLOBAL TOTAL** | **[units]** | **$[avg]** | **$[amount]** | **100%** | **$[amount]** |
"""

GEOGRAPHIC_BREAKDOWN_REGIONAL = """
**REGIONAL MARKET BREAKDOWN ({geographic_scope}):**

For the specified geographic scope, provide sub-regional or segment breakdown as appropriate.

| Sub-Region/Segment | Population/Units | % of Regional TAM | Key Characteristics |
|--------------------|-----------------|-------------------|---------------------|
| [Sub-region 1] | [number] | [%] | [Characteristics] |
| [Sub-region 2] | [number] | [%] | [Characteristics] |
| [Sub-region 3] | [number] | [%] | [Characteristics] |
| **TOTAL** | **[number]** | **100%** | |
"""

# =============================================================================
# SEGMENTATION INSTRUCTIONS (Used by V3 template)
# =============================================================================

SEGMENTATION_PROVIDED = """
Based on the target industry "{industry}" and the innovation description, analyze value creation across the following market segments. If specific segments were not provided, identify the 3-5 most relevant segments for this innovation.
"""

SEGMENTATION_NOT_PROVIDED = """
**IMPORTANT:** No specific market segmentation was provided. Based on the innovation description and target industry, identify **3-5 key market segments** that would benefit from this innovation. For each segment:

1. Define the segment clearly (who they are, what characterizes them)
2. Estimate the segment size as % of total TAM
3. Assess value creation potential for each segment
4. Rank segments by value density (value per unit)

Consider segmentation dimensions such as:
- Clinical/use case (e.g., post-surgery vs. trauma vs. chronic monitoring)
- Care setting (e.g., ICU, step-down, ambulatory, home)
- Patient risk level (e.g., high-risk, moderate-risk)
- Institution type (e.g., academic medical center, community hospital, clinic)
- Payer type (e.g., commercial, Medicare, Medicaid, self-pay)
"""

# =============================================================================
# VALUE CREATION MODEL TEMPLATE V3 - ENHANCED WITH GROWTH RATES
# =============================================================================

VALUE_CREATION_TEMPLATE_V3 = """You are acting as an expert business innovation economist, financial modeler, and {industry} sector specialist with deep experience in value quantification for venture capital and commercialization purposes.

I need you to analyze the following innovation and develop a comprehensive value creation model.

═══════════════════════════════════════════════════════════════════════════════
ANALYSIS OBJECTIVE
═══════════════════════════════════════════════════════════════════════════════

**Primary Goal:** Determine the TOTAL POTENTIAL VALUE CREATION for this innovation, assuming 100% market penetration within the defined geographic scope and market segment(s).

**Key Deliverables:**
1. Total value creation at full market penetration (base year 2026)
2. Value breakdown by geography (if Global scope)
3. Value breakdown by market segment
4. Monetization potential for each value driver and stakeholder
5. **Annual growth rates for all key value-driving parameters**
6. **Multi-year projection model from 2026 to 2040**
7. Excel-ready financial model with linked formulas

═══════════════════════════════════════════════════════════════════════════════
INNOVATION OVERVIEW
═══════════════════════════════════════════════════════════════════════════════

**Innovation Name:** {innovation_name}

**Target Market/Industry:** {industry}

**Geographic Scope:** {geographic_scope}

**Analysis Timeframe:** {analysis_timeframe}

**Innovation Stage:** {innovation_stage}

───────────────────────────────────────────────────────────────────────────────
DETAILED DESCRIPTION
───────────────────────────────────────────────────────────────────────────────

{innovation_description}

{problem_section}

{customer_section}

{advantage_section}

───────────────────────────────────────────────────────────────────────────────
MARKET CONTEXT (User-Provided Estimates)
───────────────────────────────────────────────────────────────────────────────

{market_section}

───────────────────────────────────────────────────────────────────────────────
REGULATORY & IP CONTEXT
───────────────────────────────────────────────────────────────────────────────

{regulatory_section}

───────────────────────────────────────────────────────────────────────────────
KEY CONSIDERATIONS
───────────────────────────────────────────────────────────────────────────────

{risks_section}

{assumptions_section}

═══════════════════════════════════════════════════════════════════════════════
ANALYSIS REQUIREMENTS
═══════════════════════════════════════════════════════════════════════════════

Please provide the following analysis structured for input into an Excel financial model.
All calculations should assume **100% market penetration** to establish maximum value potential.
**Base year for all calculations is 2026.**

───────────────────────────────────────────────────────────────────────────────
SECTION 1: TOTAL ADDRESSABLE MARKET (TAM) AT 100% PENETRATION
───────────────────────────────────────────────────────────────────────────────

Define the complete market universe assuming full adoption.
**For each parameter, provide the 2026 base value AND the estimated annual rate of change.**

| Parameter | 2026 Value | Unit | Annual Growth Rate (%) | Growth Rationale |
|-----------|------------|------|----------------------|------------------|
| Total Addressable Population/Units | | | [+X% or -X%] | [Why this growth rate] |
| Eligible/Qualified Market (% of TAM) | | % | [+X% or -X%] | [Regulatory/clinical changes] |
| Total Serviceable Market | | (calculated) | (derived) | |
| Average Annual Usage per Unit | | [unit] | [+X% or -X%] | [Usage trend drivers] |
| Total Annual Usage Units (100% pen.) | | (calculated) | (derived) | |
| Current Baseline Performance | | [metric] | [+X% or -X%] | [Technology improvement] |
| Innovation Performance | | [metric] | [+X% or -X%] | [R&D/iteration impact] |
| Performance Improvement Delta | | % or absolute | (derived) | |

**Key Market Growth Drivers:**
1. [Driver 1] — impact on growth rate
2. [Driver 2] — impact on growth rate
3. [Driver 3] — impact on growth rate

{geographic_breakdown_section}

───────────────────────────────────────────────────────────────────────────────
SECTION 2: MARKET SEGMENTATION ANALYSIS
───────────────────────────────────────────────────────────────────────────────

{segmentation_instructions}

**SEGMENT ANALYSIS TABLE (with Growth Rates):**

| Segment | Description | 2026 TAM Size | % of Total | Annual Growth Rate (%) | Growth Rationale |
|---------|-------------|---------------|------------|----------------------|------------------|
| Segment 1 | [Description] | [units] | [%] | [+X%] | [Why this segment grows at this rate] |
| Segment 2 | [Description] | [units] | [%] | [+X%] | [Why this segment grows at this rate] |
| Segment 3 | [Description] | [units] | [%] | [+X%] | [Why this segment grows at this rate] |
| Segment 4 | [Description] | [units] | [%] | [+X%] | [Why this segment grows at this rate] |
| Segment 5 | [Description] | [units] | [%] | [+X%] | [Why this segment grows at this rate] |
| **TOTAL** | | | **100%** | **[weighted avg]** | |

───────────────────────────────────────────────────────────────────────────────
SECTION 3: STAKEHOLDER IDENTIFICATION
───────────────────────────────────────────────────────────────────────────────

Identify exactly 4 key stakeholders who will derive value from this innovation.
For each stakeholder, provide:

| ID | Stakeholder Name | Type | Role in Value Chain | Value Capture Mechanism |
|----|------------------|------|---------------------|------------------------|
| S1 | [Primary buyer/customer] | Customer | [How they purchase/adopt] | [How they capture value] |
| S2 | [End user/beneficiary] | End User | [How they benefit] | [Direct benefit type] |
| S3 | [Partner/ecosystem player] | Partner | [Their role in ecosystem] | [Partnership value] |
| S4 | [Broader beneficiary] | Society/System | [Indirect benefits] | [Externality capture] |

*Stakeholder Types: Customer, End User, Partner, Internal, Society, Payer, Regulator*
{customer_note}

───────────────────────────────────────────────────────────────────────────────
SECTION 4: VALUE DRIVER ANALYSIS (Top 5 Drivers) — WITH GROWTH RATES
───────────────────────────────────────────────────────────────────────────────

For EACH of the 5 value drivers, provide complete specifications at FULL MARKET PENETRATION.
**Include annual growth rate estimates for each calculation factor.**

**DRIVER [D1]: [Driver Name]**

| Attribute | Specification |
|-----------|---------------|
| **Driver Name** | [Clear, descriptive name] |
| **Category** | [Cost Reduction / Revenue Enhancement / Risk Mitigation / Strategic Value / Productivity Gain] |
| **Unit of Measurement** | [e.g., per patient, per procedure, per device, per year] |
| **Business Rationale** | [2-3 sentences explaining WHY this creates value] |

**Calculation Factors (at 100% penetration) — WITH GROWTH RATES:**

| Factor | Name | 2026 Value | Unit | Annual Growth Rate (%) | Growth Rationale |
|--------|------|------------|------|----------------------|------------------|
| Factor 1 | [Primary quantity] | [number] | [unit] | [+X% or -X%] | [Why this factor changes] |
| Factor 2 | [Rate/price/multiplier] | [number] | [unit] | [+X% or -X%] | [Price/cost trends] |
| Factor 3 | [Probability/efficiency] | [0-1] | factor | [+X% or -X%] | [Improvement trajectory] |

**Volume at 100% Penetration:**

| Attribute | 2026 Value | Annual Growth Rate (%) | Derivation |
|-----------|------------|----------------------|------------|
| Total Annual Units | [number] | [+X%] | [From Section 1 TAM] |

**2026 Value Calculation (100% Penetration):**
- Value Per Unit = Factor 1 × Factor 2 × Factor 3 = $[amount]
- **Total Annual Value (2026, 100% Pen.)** = Value Per Unit × Total Units = **$[amount]**

**Projected Annual Growth Rate for This Driver:** [+X% or -X%]
(Composite of factor growth rates)

[Repeat for D2, D3, D4, D5]

───────────────────────────────────────────────────────────────────────────────
SECTION 5: STAKEHOLDER VALUE ALLOCATION & MONETIZATION POTENTIAL
───────────────────────────────────────────────────────────────────────────────

**5A. Value Allocation Matrix (% of value flowing to each stakeholder):**

| Driver | S1 (%) | S2 (%) | S3 (%) | S4 (%) | Allocation Rationale |
|--------|--------|--------|--------|--------|---------------------|
| D1 | | | | | [Why value splits this way] |
| D2 | | | | | [Why value splits this way] |
| D3 | | | | | [Why value splits this way] |
| D4 | | | | | [Why value splits this way] |
| D5 | | | | | [Why value splits this way] |

**5B. MONETIZATION POTENTIAL ANALYSIS:**

**Monetization % by Value Driver:**

| Driver | 2026 Total Value | Monetizable % | Monetizable Value | Annual Change in Monetization % | Non-Monetizable Rationale |
|--------|------------------|---------------|-------------------|-------------------------------|--------------------------|
| D1: [name] | $[amount] | [%] | $[amount] | [+X% or 0%] | [Why remainder is not monetizable] |
| D2: [name] | $[amount] | [%] | $[amount] | [+X% or 0%] | [Why remainder is not monetizable] |
| D3: [name] | $[amount] | [%] | $[amount] | [+X% or 0%] | [Why remainder is not monetizable] |
| D4: [name] | $[amount] | [%] | $[amount] | [+X% or 0%] | [Why remainder is not monetizable] |
| D5: [name] | $[amount] | [%] | $[amount] | [+X% or 0%] | [Why remainder is not monetizable] |
| **TOTAL** | **$[amount]** | **[weighted %]** | **$[amount]** | | |

**Monetization % by Stakeholder:**

| Stakeholder | 2026 Value Received | Monetizable % | Monetizable Value | Monetization Mechanism |
|-------------|---------------------|---------------|-------------------|----------------------|
| S1: [name] | $[amount] | [%] | $[amount] | [How to capture] |
| S2: [name] | $[amount] | [%] | $[amount] | [How to capture] |
| S3: [name] | $[amount] | [%] | $[amount] | [How to capture] |
| S4: [name] | $[amount] | [%] | $[amount] | [How to capture] |
| **TOTAL** | **$[amount]** | **[weighted %]** | **$[amount]** | |

───────────────────────────────────────────────────────────────────────────────
SECTION 6: VALUE SUMMARY — 2026 BASE YEAR (100% PENETRATION)
───────────────────────────────────────────────────────────────────────────────

**6A. Total Value Creation (2026 at 100% Penetration):**

| Driver | 2026 Annual Value | % of Total | Projected Growth Rate |
|--------|-------------------|------------|----------------------|
| D1: [name] | $[amount] | [%] | [+X%/year] |
| D2: [name] | $[amount] | [%] | [+X%/year] |
| D3: [name] | $[amount] | [%] | [+X%/year] |
| D4: [name] | $[amount] | [%] | [+X%/year] |
| D5: [name] | $[amount] | [%] | [+X%/year] |
| **TOTAL VALUE CREATION** | **$[amount]** | **100%** | **[weighted avg]** |

**6B. Value by Stakeholder (2026 at 100% Penetration):**

| Stakeholder | 2026 Total Value | % of Total | Monetizable Value | Monetizable % |
|-------------|------------------|------------|-------------------|---------------|
| S1: [name] | $[amount] | [%] | $[amount] | [%] |
| S2: [name] | $[amount] | [%] | $[amount] | [%] |
| S3: [name] | $[amount] | [%] | $[amount] | [%] |
| S4: [name] | $[amount] | [%] | $[amount] | [%] |
| **TOTAL** | **$[amount]** | **100%** | **$[amount]** | **[%]** |

{geographic_value_section}

**6D. Value by Market Segment (2026):**

| Segment | 2026 Total Value | % of Total | Growth Rate | Primary Beneficiary |
|---------|------------------|------------|-------------|---------------------|
| Segment 1: [name] | $[amount] | [%] | [+X%/year] | [Stakeholder] |
| Segment 2: [name] | $[amount] | [%] | [+X%/year] | [Stakeholder] |
| Segment 3: [name] | $[amount] | [%] | [+X%/year] | [Stakeholder] |
| Segment 4: [name] | $[amount] | [%] | [+X%/year] | [Stakeholder] |
| Segment 5: [name] | $[amount] | [%] | [+X%/year] | [Stakeholder] |
| **TOTAL** | **$[amount]** | **100%** | **[weighted]** | |

───────────────────────────────────────────────────────────────────────────────
SECTION 7: GROWTH RATE SUMMARY — ALL KEY PARAMETERS
───────────────────────────────────────────────────────────────────────────────

**Comprehensive Growth Rate Table for Projection Model:**

| Parameter Category | Parameter | 2026 Base Value | Annual Growth Rate (%) | 2040 Projected Value | Growth Driver |
|-------------------|-----------|-----------------|----------------------|---------------------|---------------|
| **Market Size** | | | | | |
| | Total Addressable Market | [value] | [+X%] | [calculated] | [driver] |
| | Eligible Market % | [value] | [+X%] | [calculated] | [driver] |
| | Market Penetration Curve | [value] | [varies] | [calculated] | [adoption model] |
| **Value Drivers** | | | | | |
| | D1: [name] - Factor 1 | [value] | [+X%] | [calculated] | [driver] |
| | D1: [name] - Factor 2 | [value] | [+X%] | [calculated] | [driver] |
| | D1: [name] - Factor 3 | [value] | [+X%] | [calculated] | [driver] |
| | D2: [name] - Factor 1 | [value] | [+X%] | [calculated] | [driver] |
| | [Continue for all drivers...] | | | | |
| **Pricing/Costs** | | | | | |
| | Average Price/Unit | [value] | [+X%] | [calculated] | [market dynamics] |
| | Cost per Intervention | [value] | [+X%] | [calculated] | [healthcare inflation] |
| **Monetization** | | | | | |
| | Overall Monetization % | [value] | [+X%] | [calculated] | [market maturity] |
| **External Factors** | | | | | |
| | Healthcare Inflation | N/A | [+X%] | N/A | [economic baseline] |
| | Technology Cost Deflation | N/A | [-X%] | N/A | [Moore's law equivalent] |
| | Regulatory Expansion | N/A | [+X%] | N/A | [policy trends] |

**Growth Rate Categorization:**
- **High Growth (>10%/year):** [List parameters]
- **Moderate Growth (3-10%/year):** [List parameters]
- **Low/Stable Growth (0-3%/year):** [List parameters]
- **Declining (-X%/year):** [List parameters]

───────────────────────────────────────────────────────────────────────────────
SECTION 8: MULTI-YEAR VALUE PROJECTION (2026-2040)
───────────────────────────────────────────────────────────────────────────────

**8A. Total Value Creation Projection (100% Penetration Scenario):**

| Year | Total Value Creation | YoY Growth | Cumulative Value | Monetizable Value |
|------|---------------------|------------|------------------|-------------------|
| 2026 | $[amount] | — | $[amount] | $[amount] |
| 2027 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2028 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2029 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2030 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2031 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2032 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2033 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2034 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2035 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2036 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2037 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2038 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2039 | $[amount] | [+X%] | $[amount] | $[amount] |
| 2040 | $[amount] | [+X%] | $[amount] | $[amount] |

**8B. Value by Driver Over Time (Selected Years):**

| Driver | 2026 | 2030 | 2035 | 2040 | CAGR |
|--------|------|------|------|------|------|
| D1: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| D2: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| D3: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| D4: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| D5: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| **TOTAL** | **$[amount]** | **$[amount]** | **$[amount]** | **$[amount]** | **[X%]** |

**8C. Value by Stakeholder Over Time (Selected Years):**

| Stakeholder | 2026 | 2030 | 2035 | 2040 | CAGR |
|-------------|------|------|------|------|------|
| S1: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| S2: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| S3: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| S4: [name] | $[amount] | $[amount] | $[amount] | $[amount] | [X%] |
| **TOTAL** | **$[amount]** | **$[amount]** | **$[amount]** | **$[amount]** | **[X%]** |

**8D. Key Projection Milestones:**

| Milestone | Year | Value | Significance |
|-----------|------|-------|--------------|
| $1B Total Value Creation | [year] | $1.0B | [Market significance] |
| $500M Monetizable Value | [year] | $500M | [Commercial milestone] |
| Peak Growth Rate Year | [year] | [+X%] | [Inflection point] |
| Market Maturity (growth <5%) | [year] | [value] | [Steady state] |

───────────────────────────────────────────────────────────────────────────────
SECTION 9: KEY ASSUMPTIONS & SENSITIVITY
───────────────────────────────────────────────────────────────────────────────

**Critical Assumptions (Including Growth Rate Assumptions):**
{assumptions_note}

| # | Assumption | 2026 Base | Growth Rate | Conservative | Optimistic | Value Impact |
|---|------------|-----------|-------------|--------------|------------|--------------|
| 1 | [Assumption] | [value] | [+X%/yr] | [-Y%/yr] | [+Z%/yr] | ±$[impact] |
| 2 | [Assumption] | [value] | [+X%/yr] | [-Y%/yr] | [+Z%/yr] | ±$[impact] |
| 3 | [Assumption] | [value] | [+X%/yr] | [-Y%/yr] | [+Z%/yr] | ±$[impact] |
| 4 | [Assumption] | [value] | [+X%/yr] | [-Y%/yr] | [+Z%/yr] | ±$[impact] |
| 5 | [Assumption] | [value] | [+X%/yr] | [-Y%/yr] | [+Z%/yr] | ±$[impact] |

**Growth Rate Sensitivity (Impact on 2040 Value):**

| Parameter | Base Growth | -2% Growth | +2% Growth | 2040 Value Impact |
|-----------|-------------|------------|------------|-------------------|
| TAM Growth | [X%] | [X-2%] | [X+2%] | ±$[amount] (±[%]) |
| Price/Value per Unit | [X%] | [X-2%] | [X+2%] | ±$[amount] (±[%]) |
| Monetization % | [X%] | [X-2%] | [X+2%] | ±$[amount] (±[%]) |

───────────────────────────────────────────────────────────────────────────────
SECTION 10: COMMERCIALIZATION METRICS
───────────────────────────────────────────────────────────────────────────────

| Metric | 2026 Value | 2030 Value | 2035 Value | 2040 Value |
|--------|------------|------------|------------|------------|
| **Total Annual Value Creation** | **$[amount]** | **$[amount]** | **$[amount]** | **$[amount]** |
| Total Monetizable Value | $[amount] | $[amount] | $[amount] | $[amount] |
| Value per Unit | $[amount] | $[amount] | $[amount] | $[amount] |
| Monetizable Value per Unit | $[amount] | $[amount] | $[amount] | $[amount] |
| **15-Year Cumulative Value (2026-2040)** | | | | **$[total]** |
| **15-Year Cumulative Monetizable** | | | | **$[total]** |

**Pricing Guidance (Based on 2026 Monetizable Value):**
| Price Point | Calculation | Annual Revenue Potential (100% pen.) |
|-------------|-------------|-------------------------------------|
| Conservative (10% capture) | $[amount] | $[amount] |
| Moderate (20% capture) | $[amount] | $[amount] |
| Aggressive (30% capture) | $[amount] | $[amount] |

───────────────────────────────────────────────────────────────────────────────
SECTION 11: RISK-ADJUSTED VALUE ASSESSMENT
───────────────────────────────────────────────────────────────────────────────

{risks_assessment}

| Risk Category | Specific Risk | Probability | Impact on Growth Rate | Value at Risk (2040) | Mitigation |
|---------------|--------------|-------------|----------------------|---------------------|------------|
| Technical | [Risk] | [%] | [Reduces growth by X%] | $[amount] | [Strategy] |
| Market/Adoption | [Risk] | [%] | [Reduces growth by X%] | $[amount] | [Strategy] |
| Regulatory | [Risk] | [%] | [Reduces growth by X%] | $[amount] | [Strategy] |
| Competitive | [Risk] | [%] | [Reduces growth by X%] | $[amount] | [Strategy] |
| Execution | [Risk] | [%] | [Reduces growth by X%] | $[amount] | [Strategy] |

**Risk-Adjusted Value Summary:**
- 2026 Base Case Value: $[amount]
- 2040 Base Case Value: $[amount]
- Risk Adjustment Factor: [0.X]
- 2040 Risk-Adjusted Value: $[amount]

═══════════════════════════════════════════════════════════════════════════════
OUTPUT CONSTRAINTS
═══════════════════════════════════════════════════════════════════════════════

1. All calculations assume **100% market penetration** to show maximum value potential
2. **Base year is 2026** for all values
3. **Provide annual growth rates** for every key value-driving parameter
4. Use conservative estimates; cite ranges where uncertain (use midpoint)
5. All monetary values in {currency}
6. Clearly distinguish between INPUT values, GROWTH RATES, and CALCULATED values
7. Flag high-uncertainty growth rate assumptions with [⚠️ HIGH UNCERTAINTY]
8. Ensure all allocations sum to 100%
9. Format numbers for Excel import (no text in numeric fields)

**IMPORTANT: Do NOT include a bibliography, references, sources section, or citations list at the end of your response. All rationale should be provided inline within the analysis tables.**

═══════════════════════════════════════════════════════════════════════════════
EXCEL MODEL IMPLEMENTATION
═══════════════════════════════════════════════════════════════════════════════

After providing the analysis above, generate a Python script to create an Excel financial model.

**WORKBOOK STRUCTURE (7 sheets required):**
1. **Dashboard** — Executive summary with KPIs, total value, key charts, 2026 vs 2040 comparison
2. **TAM Analysis** — Market sizing with geographic/segment breakdowns and growth rates
3. **Value Drivers** — Detailed calculations for each driver with growth rate inputs
4. **Stakeholders** — Value allocation and monetization analysis
5. **Growth Rates** — Consolidated table of all growth rate assumptions (editable inputs)
6. **Projections (2026-2040)** — Year-by-year projection model with charts
7. **Sensitivity** — Scenario analysis with growth rate variations

═══════════════════════════════════════════════════════════════════════════════
STRICT TECHNICAL REQUIREMENTS
═══════════════════════════════════════════════════════════════════════════════

**CRITICAL RULES:**
• Workbook variable MUST be named exactly: wb
• Do NOT include wb.save() or wb.close() — the application handles file saving
• Use only openpyxl library (do NOT use pandas, xlsxwriter, or other libraries)
• All numeric cells should contain formulas referencing other cells where applicable
• **Growth rates should be input cells that drive all projection calculations**
• **Projections sheet must have formulas that reference growth rates (not hardcoded values)**

───────────────────────────────────────────────────────────────────────────────
AVAILABLE IMPORTS & OBJECTS
───────────────────────────────────────────────────────────────────────────────

**Core openpyxl:**
• openpyxl, Workbook

**Styling:**
• Font — text formatting (bold, color, size, italic, underline)
• PatternFill — cell background colors (solid, pattern fills)
• GradientFill — gradient backgrounds
• Alignment — text alignment (horizontal, vertical, wrap_text, rotation)
• Border, Side — cell borders (thin, medium, thick, double, dashed)
• Color — color definitions
• Protection — cell/sheet protection
• NamedStyle — reusable named styles

**Number Formats:**
• Currency: '"$"#,##0.00' or '$#,##0' or '_($* #,##0.00_)'
• Percentage: '0%' or '0.00%'
• Number: '#,##0' or '#,##0.00'
• Millions: '#,##0,,"M"'
• Billions: '#,##0,,,"B"'

**Charts (all types supported):**
• BarChart, BarChart3D — vertical/horizontal bar charts
• LineChart, LineChart3D — line graphs (ideal for projections over time)
• AreaChart, AreaChart3D — stacked area charts (good for cumulative value)
• PieChart, PieChart3D — pie charts with data labels
• DoughnutChart — ring charts
• ScatterChart — XY scatter plots
• BubbleChart — bubble charts
• RadarChart — spider/radar charts
• StockChart — OHLC stock price charts
• SurfaceChart, SurfaceChart3D — 3D surface plots

**Chart Components:**
• Reference — defines data ranges for charts
• Series — individual data series
• DataLabelList — configure data labels
• Legend — chart legend configuration

**Conditional Formatting:**
• ColorScaleRule — 2-color or 3-color gradient scales
• DataBarRule — in-cell data bars
• IconSetRule — icon sets (arrows, traffic lights, flags)
• CellIsRule — highlight cells based on value conditions
• FormulaRule — highlight based on custom formulas

**Data Validation:**
• DataValidation — dropdown lists, number ranges, custom validation

**Tables:**
• Table — structured Excel tables with sorting/filtering
• TableStyleInfo — table styling

**Comments:**
• Comment — cell comments/notes

**Utilities:**
• get_column_letter — convert column number to letter
• column_index_from_string — convert letter to number

**Date/Time:**
• datetime, date, timedelta
• relativedelta — relative date calculations

───────────────────────────────────────────────────────────────────────────────
HELPER CLASSES (Pre-built utilities)
───────────────────────────────────────────────────────────────────────────────

**StyleFactory** — Pre-built cell styles:
• StyleFactory.header_style(bg_color, font_color, bold, font_size)
• StyleFactory.subheader_style(bg_color, font_color)
• StyleFactory.data_style(align)
• StyleFactory.currency_style()
• StyleFactory.percentage_style()
• StyleFactory.title_style()
• StyleFactory.kpi_value_style(color)
• StyleFactory.highlight_positive(), .highlight_negative()
• StyleFactory.thin_border(), .thick_border()
• StyleFactory.apply_style(cell, style_dict)

**ChartFactory** — Easy chart creation:
• ChartFactory.create_bar_chart(ws, data_range, categories_range, title, position, ...)
• ChartFactory.create_line_chart(ws, data_range, categories_range, title, position, smooth, ...)
• ChartFactory.create_pie_chart(ws, data_range, categories_range, title, position, ...)
• ChartFactory.create_doughnut_chart(ws, ...)
• ChartFactory.create_scatter_chart(ws, x_range, y_range, ...)
• ChartFactory.create_area_chart(ws, data_range, categories_range, title, position, stacked, ...)
• ChartFactory.create_radar_chart(ws, ...)

**FormulaHelper** — Excel formula generators:
• FormulaHelper.sum(range), .average(range), .count(range)
• FormulaHelper.min_val(range), .max_val(range)
• FormulaHelper.countif(range, criteria), .sumif(range, criteria, sum_range)
• FormulaHelper.vlookup(value, table, col, exact), .index_match(...)
• FormulaHelper.if_formula(condition, true_val, false_val)
• FormulaHelper.iferror(formula, error_value)
• FormulaHelper.npv(rate, values), .irr(values), .xirr(values, dates)
• FormulaHelper.pmt(rate, nper, pv), .pv(...), .fv(...)
• FormulaHelper.cagr(start, end, periods)
• FormulaHelper.growth_rate(new, old)

**BusinessModelBuilder** — Business model components:
• BusinessModelBuilder.create_kpi_card(ws, row, col, title, value, subtitle, color)
• BusinessModelBuilder.create_data_table(ws, row, col, headers, data, table_name, style)
• BusinessModelBuilder.create_scenario_table(ws, row, col, scenarios_dict, metrics_list)
• BusinessModelBuilder.create_assumption_log(ws, row, col, assumptions_list)
• BusinessModelBuilder.create_waterfall_data(ws, row, col, items_list, title)

**ConditionalFormatHelper** — Conditional formatting:
• ConditionalFormatHelper.add_color_scale(ws, range, start_color, mid_color, end_color)
• ConditionalFormatHelper.add_data_bars(ws, range, color)
• ConditionalFormatHelper.add_icon_set(ws, range, icon_style)
• ConditionalFormatHelper.highlight_cells_greater_than(ws, range, value, fill_color)
• ConditionalFormatHelper.highlight_cells_less_than(ws, range, value, fill_color)

**DataValidationHelper** — Data validation:
• DataValidationHelper.create_dropdown(ws, range, options_list, ...)
• DataValidationHelper.create_number_range(ws, range, min_val, max_val)
• DataValidationHelper.create_percentage_validation(ws, range)

**WorksheetUtils** — Worksheet utilities:
• WorksheetUtils.auto_fit_columns(ws, min_width, max_width)
• WorksheetUtils.set_column_width(ws, column, width)
• WorksheetUtils.freeze_panes(ws, cell)
• WorksheetUtils.add_auto_filter(ws, range)
• WorksheetUtils.merge_cells(ws, range)
• WorksheetUtils.add_comment(ws, cell, text, author)
• WorksheetUtils.add_hyperlink(ws, cell, url, display_text)
• WorksheetUtils.protect_sheet(ws, password)

**TableBuilder** — Excel tables:
• TableBuilder.create_table(ws, range, name, style_name)

**Brand Colors:**
• BRAND_COLORS["primary_blue"] = "0067B9"
• BRAND_COLORS["dark_blue"] = "003D6A"
• BRAND_COLORS["light_blue"] = "4A9BD9"
• BRAND_COLORS["accent_orange"] = "FF6600"
• BRAND_COLORS["success_green"] = "28A745"
• BRAND_COLORS["error_red"] = "DC3545"
• BRAND_COLORS["warning_yellow"] = "FFC107"

───────────────────────────────────────────────────────────────────────────────
PROJECTIONS SHEET REQUIREMENTS
───────────────────────────────────────────────────────────────────────────────

The **Projections (2026-2040)** sheet must include:

1. **Year columns:** 2026, 2027, 2028, ... , 2040 (15 years)

2. **Row sections:**
   - Market Size (TAM) projection
   - Value by Driver (D1-D5) projection
   - Total Value Creation projection
   - Monetizable Value projection
   - Value by Stakeholder projection
   - Cumulative Value tracker

3. **Formula structure:**
   - 2026 column: Reference base values from Value Drivers sheet
   - 2027+ columns: = Prior Year × (1 + Growth Rate)
   - Growth rates should reference the Growth Rates sheet (not hardcoded)

4. **Charts to include:**
   - Line chart: Total Value Creation over time (2026-2040)
   - Stacked area chart: Value by Driver over time
   - Line chart: Monetizable vs Total Value over time

5. **Summary metrics:**
   - 15-Year Cumulative Value
   - CAGR (2026-2040)
   - Year when value exceeds $1B (if applicable)
"""

# =============================================================================
# EXECUTIVE PITCH DECK TEMPLATE
# =============================================================================

EXECUTIVE_PITCH_TEMPLATE = """You are acting as an expert business strategist and presentation designer specializing in executive communications for the {industry} sector.

I need you to create content for an executive pitch deck for the following innovation.

═══════════════════════════════════════════════════════════════════════════════
PRESENTATION OBJECTIVE
═══════════════════════════════════════════════════════════════════════════════

**Primary Goal:** Create a compelling executive pitch deck that clearly communicates the value proposition, market opportunity, and investment case for this innovation.

**Target Audience:** C-suite executives, board members, and investment committee

**Presentation Length:** 10-12 slides maximum

═══════════════════════════════════════════════════════════════════════════════
INNOVATION OVERVIEW
═══════════════════════════════════════════════════════════════════════════════

**Innovation Name:** {innovation_name}

**Target Market/Industry:** {industry}

**Geographic Scope:** {geographic_scope}

**Innovation Stage:** {innovation_stage}

───────────────────────────────────────────────────────────────────────────────
DETAILED DESCRIPTION
───────────────────────────────────────────────────────────────────────────────

{innovation_description}

{problem_section}

{customer_section}

{advantage_section}

───────────────────────────────────────────────────────────────────────────────
MARKET CONTEXT
───────────────────────────────────────────────────────────────────────────────

{market_section}

───────────────────────────────────────────────────────────────────────────────
REGULATORY & IP CONTEXT
───────────────────────────────────────────────────────────────────────────────

{regulatory_section}

───────────────────────────────────────────────────────────────────────────────
KEY CONSIDERATIONS
───────────────────────────────────────────────────────────────────────────────

{risks_section}

{assumptions_section}

═══════════════════════════════════════════════════════════════════════════════
SLIDE DECK STRUCTURE
═══════════════════════════════════════════════════════════════════════════════

Please create content for each of the following slides:

───────────────────────────────────────────────────────────────────────────────
SLIDE 1: TITLE SLIDE
───────────────────────────────────────────────────────────────────────────────
- Innovation name (prominent)
- Tagline (one compelling sentence)
- Presenter/Company name
- Date

───────────────────────────────────────────────────────────────────────────────
SLIDE 2: THE PROBLEM
───────────────────────────────────────────────────────────────────────────────
- 3 bullet points describing the problem
- Key statistic highlighting problem severity
- Visual suggestion (icon/image concept)
- Speaker notes (what to emphasize verbally)

───────────────────────────────────────────────────────────────────────────────
SLIDE 3: THE SOLUTION
───────────────────────────────────────────────────────────────────────────────
- Innovation description (2-3 sentences)
- 3-4 key features/capabilities
- How it solves the problem
- Visual suggestion (diagram/illustration concept)
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 4: VALUE PROPOSITION
───────────────────────────────────────────────────────────────────────────────
- Primary value statement (bold, prominent)
- 3 supporting value points
- Quantified benefit (if available)
- Differentiation statement
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 5: MARKET OPPORTUNITY
───────────────────────────────────────────────────────────────────────────────
- TAM / SAM / SOM breakdown
- Market growth rate
- Key market trends (3 bullets)
- Visual: Market sizing chart concept
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 6: COMPETITIVE LANDSCAPE
───────────────────────────────────────────────────────────────────────────────
- Key competitors (3-5)
- Competitive positioning matrix (axes to use)
- Our differentiation (3 points)
- Why we win
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 7: BUSINESS MODEL
───────────────────────────────────────────────────────────────────────────────
- Revenue model description
- Pricing strategy/range
- Unit economics (if applicable)
- Path to profitability
- Visual: Business model canvas elements
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 8: GO-TO-MARKET STRATEGY
───────────────────────────────────────────────────────────────────────────────
- Target customer segments (prioritized)
- Channel strategy
- Partnership opportunities
- Key milestones (6-18 months)
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 9: FINANCIAL PROJECTIONS
───────────────────────────────────────────────────────────────────────────────
- 5-year revenue projection (table)
- Key assumptions (3-4 bullets)
- Break-even timeline
- Visual: Revenue growth chart concept
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 10: TEAM & CAPABILITIES
───────────────────────────────────────────────────────────────────────────────
- Key team members/roles
- Relevant experience highlights
- Advisory board (if applicable)
- Core competencies
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 11: THE ASK
───────────────────────────────────────────────────────────────────────────────
- Investment/resource request
- Use of funds breakdown
- Timeline and milestones
- Expected outcomes/returns
- Speaker notes

───────────────────────────────────────────────────────────────────────────────
SLIDE 12: SUMMARY & NEXT STEPS
───────────────────────────────────────────────────────────────────────────────
- 3 key takeaways
- Why now (urgency)
- Specific call to action
- Contact information
- Speaker notes

═══════════════════════════════════════════════════════════════════════════════
OUTPUT FORMAT REQUIREMENTS
═══════════════════════════════════════════════════════════════════════════════

For each slide, provide:
1. **Slide Title**
2. **Main Content** (bullet points, key text)
3. **Visual Suggestion** (chart type, diagram, image concept)
4. **Speaker Notes** (what to say, key points to emphasize)
5. **Design Notes** (color emphasis, animation suggestions)

All monetary values in {currency}.

**IMPORTANT: Do NOT include a bibliography, references, or sources section at the end of your response.**

═══════════════════════════════════════════════════════════════════════════════
PYTHON-PPTX SCRIPT GENERATION
═══════════════════════════════════════════════════════════════════════════════

After providing the slide content above, generate a Python script using python-pptx to create the PowerPoint file.

**CRITICAL RULES:**
• Presentation variable MUST be named exactly: prs
• Do NOT include prs.save() or prs.close() — the application handles file saving
• Use only python-pptx library
• Include professional formatting and consistent styling
• Add charts where appropriate using python-pptx chart capabilities

**Available Imports:**
```python
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RgbColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
"""

# =============================================================================
# DEEP DIVE SUMMARY - RECOMMEND TO EXPLORE TEMPLATE
# =============================================================================

GONOGO_REPORT_TEMPLATE = """You are acting as a senior innovation analyst and strategic advisor at Analog Devices' Analog Garage innovation unit, specializing in technology commercialization and investment decisions for the {industry} sector.

I need you to create a comprehensive Deep Dive Summary document with a recommendation to proceed to an Exploration project for the following innovation domain.

═══════════════════════════════════════════════════════════════════════════════
DOCUMENT OBJECTIVE
═══════════════════════════════════════════════════════════════════════════════

**Primary Goal:** Produce a structured Deep Dive Summary document that provides a thorough assessment of a technology/market domain and delivers a clear "Recommend to Explore" decision with supporting rationale.

**Target Audience:** ADI Innovation leadership, Analog Garage stakeholders, investment committee

**Document Purpose:** Support the decision to advance from Deep Dive assessment to Exploration project phase

═══════════════════════════════════════════════════════════════════════════════
INNOVATION DOMAIN OVERVIEW
═══════════════════════════════════════════════════════════════════════════════

**Innovation/Domain Name:** {innovation_name}

**Target Market/Industry:** {industry}

**Geographic Scope:** {geographic_scope}

**Innovation Stage:** {innovation_stage}

**Analysis Timeframe:** {analysis_timeframe}

───────────────────────────────────────────────────────────────────────────────
DETAILED DESCRIPTION
───────────────────────────────────────────────────────────────────────────────

{innovation_description}

{problem_section}

{customer_section}

{advantage_section}

───────────────────────────────────────────────────────────────────────────────
MARKET CONTEXT
───────────────────────────────────────────────────────────────────────────────

{market_section}

───────────────────────────────────────────────────────────────────────────────
REGULATORY & IP CONTEXT
───────────────────────────────────────────────────────────────────────────────

{regulatory_section}

───────────────────────────────────────────────────────────────────────────────
KEY CONSIDERATIONS
───────────────────────────────────────────────────────────────────────────────

{risks_section}

{assumptions_section}

═══════════════════════════════════════════════════════════════════════════════
DEEP DIVE SUMMARY DOCUMENT STRUCTURE
═══════════════════════════════════════════════════════════════════════════════

Please provide a comprehensive report following this EXACT structure:

───────────────────────────────────────────────────────────────────────────────
SECTION: EXECUTIVE SUMMARY
───────────────────────────────────────────────────────────────────────────────

**Brief Overview:**
(One paragraph summary of the domain studied, the main findings, and the recommendation to proceed to an Exploration project. What is the clearest articulation we have right now of the problem statement and why we believe it's a good fit for ADI?)

[Provide a concise, compelling paragraph that captures:
- The domain/technology area studied
- Key findings from the deep dive
- Why this represents a strategic opportunity for ADI
- The recommended path forward]

**Decision Statement:**
(Clear articulation of the "go" decision and its rationale in plain language)

[Provide a clear, direct statement such as:
"We recommend proceeding to an Exploration project for [innovation name] based on [2-3 key reasons]. This domain aligns with ADI's strategic priorities in [area] and presents a compelling opportunity to [value proposition]."]

───────────────────────────────────────────────────────────────────────────────
SECTION: PURPOSE & SCOPE
───────────────────────────────────────────────────────────────────────────────

**Objective:**
What was the goal of this assessment? (e.g., evaluate potential for entry, partnership, or investment)

[Describe the specific objectives of this Deep Dive assessment]

**Exploration Focus:**
What specific questions or hypotheses will the Exploration phase address?

[List 3-5 key questions/hypotheses for the next phase, formatted as:]
1. [Question/Hypothesis 1]
2. [Question/Hypothesis 2]
3. [Question/Hypothesis 3]
4. [Question/Hypothesis 4]
5. [Question/Hypothesis 5]

**Scope:**

*What specific subdomains, technologies, or market segments were included/excluded in the study:*
[Describe the boundaries of the Deep Dive analysis]

*What is in-scope for the Exploration phase:*
- [In-scope item 1]
- [In-scope item 2]
- [In-scope item 3]

*What is explicitly out-of-scope at this stage and why:*
- [Out-of-scope item 1] — [Reason]
- [Out-of-scope item 2] — [Reason]
- [Out-of-scope item 3] — [Reason]

**Assumptions & Constraints:**
Any key assumptions (e.g., data availability, access to testbeds), or constraints (budget, timeline, regulatory)

| Category | Assumption/Constraint | Impact on Exploration |
|----------|----------------------|----------------------|
| Data | [Assumption] | [Impact] |
| Access | [Assumption] | [Impact] |
| Budget | [Constraint] | [Impact] |
| Timeline | [Constraint] | [Impact] |
| Regulatory | [Assumption] | [Impact] |

───────────────────────────────────────────────────────────────────────────────
SECTION: CURRENT STATE OF TECHNOLOGY & MARKET
───────────────────────────────────────────────────────────────────────────────

**Technology Overview:**
Concise summary of the current technological landscape (key technologies, maturity, adoption status) — where will exploration focus to drill deeper? PoC scoping?

[Provide a structured overview including:]

*Key Technologies:*
| Technology | Maturity Level | Adoption Status | ADI Relevance |
|------------|---------------|-----------------|---------------|
| [Tech 1] | [TRL X] | [Early/Growing/Mature] | [High/Med/Low] |
| [Tech 2] | [TRL X] | [Early/Growing/Mature] | [High/Med/Low] |
| [Tech 3] | [TRL X] | [Early/Growing/Mature] | [High/Med/Low] |

*Exploration Focus Areas (Technology):*
- [Area 1 for deeper investigation]
- [Area 2 for deeper investigation]

*Potential PoC Scope:*
[Brief description of what a proof-of-concept might look like]

**Market Overview:**
Market size, growth trends, major customer segments, and relevant regulatory or economic factors — what are the key questions for exploration, around Proof of Need (PoN)? What problem statements were identified in this study (including but not limited to the one recommended to pursue in exploration)?

*Market Size & Growth:*
| Metric | Value | Growth Rate | Confidence |
|--------|-------|-------------|------------|
| TAM | $[X]B | [X]% CAGR | [High/Med/Low] |
| SAM | $[X]B | [X]% CAGR | [High/Med/Low] |
| SOM (Target) | $[X]M | [X]% CAGR | [High/Med/Low] |

*Major Customer Segments:*
1. [Segment 1] — [Size/Characteristics]
2. [Segment 2] — [Size/Characteristics]
3. [Segment 3] — [Size/Characteristics]

*Regulatory/Economic Factors:*
- [Factor 1]
- [Factor 2]
- [Factor 3]

*Key Questions for Exploration (Proof of Need):*
1. [PoN Question 1]
2. [PoN Question 2]
3. [PoN Question 3]

*Problem Statements Identified:*
| ID | Problem Statement | Priority | Recommended for Exploration |
|----|------------------|----------|----------------------------|
| P1 | [Problem] | [High/Med/Low] | [Yes — Primary Focus] |
| P2 | [Problem] | [High/Med/Low] | [No — Future consideration] |
| P3 | [Problem] | [High/Med/Low] | [No — Out of scope] |

**Key Players:**
Table or bullet list of major companies, startups, or research groups active in this space and relevance

| Player | Type | Focus Area | Relevance to ADI | Watch Priority |
|--------|------|------------|-----------------|----------------|
| [Company 1] | Incumbent | [Area] | [Relevance] | [High/Med/Low] |
| [Company 2] | Startup | [Area] | [Relevance] | [High/Med/Low] |
| [Company 3] | Research | [Area] | [Relevance] | [High/Med/Low] |
| [Company 4] | Incumbent | [Area] | [Relevance] | [High/Med/Low] |
| [Company 5] | Startup | [Area] | [Relevance] | [High/Med/Low] |

───────────────────────────────────────────────────────────────────────────────
SECTION: COMPETITIVE LANDSCAPE
───────────────────────────────────────────────────────────────────────────────

**Competitor Summary:**
Brief notes on offerings and market positions, recent moves, flag ones to watch

| Competitor | Offering | Market Position | Recent Moves | Watch Status |
|------------|----------|-----------------|--------------|--------------|
| [Competitor 1] | [Products/Services] | [Leader/Challenger/Niche] | [Recent activity] | [⚠️ Watch closely] |
| [Competitor 2] | [Products/Services] | [Leader/Challenger/Niche] | [Recent activity] | [Monitor] |
| [Competitor 3] | [Products/Services] | [Leader/Challenger/Niche] | [Recent activity] | [Monitor] |
| [Competitor 4] | [Products/Services] | [Leader/Challenger/Niche] | [Recent activity] | [⚠️ Watch closely] |

**Competitive Analysis:**
High-level SWOT or perceptual map (optional: include as a table or simple chart)

*ADI/Analog Garage Position — SWOT:*

| Strengths | Weaknesses |
|-----------|------------|
| • [Strength 1] | • [Weakness 1] |
| • [Strength 2] | • [Weakness 2] |
| • [Strength 3] | • [Weakness 3] |

| Opportunities | Threats |
|--------------|---------|
| • [Opportunity 1] | • [Threat 1] |
| • [Opportunity 2] | • [Threat 2] |
| • [Opportunity 3] | • [Threat 3] |

*Competitive Positioning Summary:*
[2-3 sentences describing ADI's potential competitive position and differentiation strategy]

**Emerging Players:**
Notable startups or new entrants to watch

| Company | Founded | Funding | Focus | Why Watch |
|---------|---------|---------|-------|-----------|
| [Startup 1] | [Year] | $[X]M | [Focus area] | [Reason] |
| [Startup 2] | [Year] | $[X]M | [Focus area] | [Reason] |
| [Startup 3] | [Year] | $[X]M | [Focus area] | [Reason] |

───────────────────────────────────────────────────────────────────────────────
SECTION: RATIONALE FOR GO DECISION
───────────────────────────────────────────────────────────────────────────────

**Key Assessment Findings:**

[Summarize the most important findings from the Deep Dive]

1. [Finding 1]
2. [Finding 2]
3. [Finding 3]
4. [Finding 4]
5. [Finding 5]

**Summary of Main Reasons to Proceed:**
(e.g., market opportunity, technology readiness, strategic fit, unique value proposition)

| Reason | Description | Confidence |
|--------|-------------|------------|
| Market Opportunity | [Description of market opportunity] | [High/Med/Low] |
| Technology Readiness | [Description of tech readiness] | [High/Med/Low] |
| Strategic Fit | [Description of strategic alignment] | [High/Med/Low] |
| Unique Value Proposition | [Description of differentiation] | [High/Med/Low] |
| Timing | [Description of market timing] | [High/Med/Low] |

**Supporting Data:**
Any critical data points, trends, or analysis that support the decision

| Data Point | Value | Significance |
|------------|-------|--------------|
| [Metric 1] | [Value] | [Why this supports the decision] |
| [Metric 2] | [Value] | [Why this supports the decision] |
| [Metric 3] | [Value] | [Why this supports the decision] |
| [Metric 4] | [Value] | [Why this supports the decision] |

───────────────────────────────────────────────────────────────────────────────
SECTION: RISKS & UNKNOWNS
───────────────────────────────────────────────────────────────────────────────

**Key Risks:**
Initial identification of major risks and uncertainties

| Risk ID | Risk Description | Category | Probability | Impact | Risk Level |
|---------|-----------------|----------|-------------|--------|------------|
| R1 | [Risk description] | [Tech/Market/Regulatory/Competitive] | [High/Med/Low] | [High/Med/Low] | [Critical/High/Medium/Low] |
| R2 | [Risk description] | [Tech/Market/Regulatory/Competitive] | [High/Med/Low] | [High/Med/Low] | [Critical/High/Medium/Low] |
| R3 | [Risk description] | [Tech/Market/Regulatory/Competitive] | [High/Med/Low] | [High/Med/Low] | [Critical/High/Medium/Low] |
| R4 | [Risk description] | [Tech/Market/Regulatory/Competitive] | [High/Med/Low] | [High/Med/Low] | [Critical/High/Medium/Low] |
| R5 | [Risk description] | [Tech/Market/Regulatory/Competitive] | [High/Med/Low] | [High/Med/Low] | [Critical/High/Medium/Low] |

**Mitigation Strategies:**
Early ideas for how these will be addressed in Exploration

| Risk ID | Mitigation Strategy | Owner | Timeline |
|---------|--------------------|----- -|----------|
| R1 | [Mitigation approach] | [Role/Team] | [When to address] |
| R2 | [Mitigation approach] | [Role/Team] | [When to address] |
| R3 | [Mitigation approach] | [Role/Team] | [When to address] |
| R4 | [Mitigation approach] | [Role/Team] | [When to address] |
| R5 | [Mitigation approach] | [Role/Team] | [When to address] |

**IP Strategy / Recommendation:**

[Provide IP strategy assessment including:]

*Current IP Landscape:*
- [Description of existing patents/IP in the space]
- [Key IP holders]
- [Freedom to operate assessment]

*Recommended IP Approach:*
- [Patent strategy recommendation]
- [Trade secret considerations]
- [Licensing opportunities/requirements]

*Note: Changes to these IP assumptions could invalidate this recommendation*

───────────────────────────────────────────────────────────────────────────────
SECTION: TECHNOLOGY TRIGGERS
───────────────────────────────────────────────────────────────────────────────

*If these change, they would invalidate or significantly alter the recommendation*

**Breakthroughs/Standards:**
What new technologies or standards would accelerate or change the opportunity?

| Trigger | Type | Current Status | If This Happens... | Impact on Recommendation |
|---------|------|----------------|-------------------|-------------------------|
| [Technology/Standard 1] | [Breakthrough/Standard] | [Current state] | [What changes] | [Accelerate/Pivot/Pause] |
| [Technology/Standard 2] | [Breakthrough/Standard] | [Current state] | [What changes] | [Accelerate/Pivot/Pause] |
| [Technology/Standard 3] | [Breakthrough/Standard] | [Current state] | [What changes] | [Accelerate/Pivot/Pause] |

───────────────────────────────────────────────────────────────────────────────
SECTION: MARKET DYNAMICS
───────────────────────────────────────────────────────────────────────────────

**Customer/Regulatory Shifts:**
What changes in demand, regulation, or competition would impact the project?

| Dynamic | Type | Current Trajectory | Trigger Scenario | Impact |
|---------|------|-------------------|------------------|--------|
| [Market dynamic 1] | [Customer/Regulatory/Competitive] | [Current trend] | [What would change] | [Positive/Negative] |
| [Market dynamic 2] | [Customer/Regulatory/Competitive] | [Current trend] | [What would change] | [Positive/Negative] |
| [Market dynamic 3] | [Customer/Regulatory/Competitive] | [Current trend] | [What would change] | [Positive/Negative] |
| [Market dynamic 4] | [Customer/Regulatory/Competitive] | [Current trend] | [What would change] | [Positive/Negative] |

───────────────────────────────────────────────────────────────────────────────
SECTION: COLLABORATION OPPORTUNITIES
───────────────────────────────────────────────────────────────────────────────

**Potential Partners/Consortia:**
Are there ecosystem changes or new partnerships that could enhance the project?

| Partner Type | Potential Partner(s) | Value to ADI | Value to Partner | Priority |
|--------------|---------------------|--------------|------------------|----------|
| Technology Partner | [Company/Institution] | [What ADI gains] | [What partner gains] | [High/Med/Low] |
| Channel Partner | [Company/Institution] | [What ADI gains] | [What partner gains] | [High/Med/Low] |
| Research Partner | [Company/Institution] | [What ADI gains] | [What partner gains] | [High/Med/Low] |
| Customer Partner | [Company/Institution] | [What ADI gains] | [What partner gains] | [High/Med/Low] |
| Consortium | [Name/Description] | [What ADI gains] | [ADI contribution] | [High/Med/Low] |

*Partnership Recommendations for Exploration Phase:*
1. [Priority partnership to pursue]
2. [Secondary partnership to explore]
3. [Consortium/ecosystem to monitor]

───────────────────────────────────────────────────────────────────────────────
SECTION: COMPANIES AND TRENDS TO WATCH
───────────────────────────────────────────────────────────────────────────────

**Key Companies:**

*Incumbents to Monitor:*
| Company | Relevance | What to Watch | Alert Trigger |
|---------|-----------|---------------|---------------|
| [Incumbent 1] | [Why relevant] | [Specific activities] | [What would trigger review] |
| [Incumbent 2] | [Why relevant] | [Specific activities] | [What would trigger review] |
| [Incumbent 3] | [Why relevant] | [Specific activities] | [What would trigger review] |

*Startups to Monitor:*
| Company | Relevance | What to Watch | Alert Trigger |
|---------|-----------|---------------|---------------|
| [Startup 1] | [Why relevant] | [Specific activities] | [What would trigger review] |
| [Startup 2] | [Why relevant] | [Specific activities] | [What would trigger review] |
| [Startup 3] | [Why relevant] | [Specific activities] | [What would trigger review] |

**Trends/Signals:**

*Notable Trends/Events to Track:*
| Trend/Signal | Type | Current Status | Why It Matters | Review Trigger |
|--------------|------|----------------|----------------|----------------|
| [Funding activity in space] | Funding | [Current state] | [Significance] | [Trigger condition] |
| [Regulatory development] | Regulatory | [Current state] | [Significance] | [Trigger condition] |
| [Technology launch] | Technology | [Current state] | [Significance] | [Trigger condition] |
| [Market consolidation] | M&A | [Current state] | [Significance] | [Trigger condition] |
| [Standard development] | Standards | [Current state] | [Significance] | [Trigger condition] |

───────────────────────────────────────────────────────────────────────────────
SECTION: FOLLOW-UP & KNOWLEDGE SHARING
───────────────────────────────────────────────────────────────────────────────

**Next Steps:**

*Immediate Actions to Launch the Exploration Project:*

| # | Action | Owner | Due Date | Status | Dependencies |
|---|--------|-------|----------|--------|--------------|
| 1 | [Action item] | [Name/Role] | [Date] | [Not Started/In Progress] | [Dependencies] |
| 2 | [Action item] | [Name/Role] | [Date] | [Not Started/In Progress] | [Dependencies] |
| 3 | [Action item] | [Name/Role] | [Date] | [Not Started/In Progress] | [Dependencies] |
| 4 | [Action item] | [Name/Role] | [Date] | [Not Started/In Progress] | [Dependencies] |
| 5 | [Action item] | [Name/Role] | [Date] | [Not Started/In Progress] | [Dependencies] |

*Known Requirements for Exploration:*
- Team: [Team composition needs]
- Budget: [Estimated budget range]
- Timeline: [Expected exploration duration]
- Resources: [Key resources needed]

**Knowledge Sharing:**

*Report Storage:*
Where will this report and future updates be stored?

| Document | Location | Access | Update Frequency |
|----------|----------|--------|------------------|
| Deep Dive Summary | [Location, e.g., SharePoint/Confluence] | [Access level] | [As needed] |
| Supporting Research | [Location] | [Access level] | [Frequency] |
| Exploration Updates | [Location] | [Access level] | [Weekly/Bi-weekly] |

**Contact:**

*Project Lead/Point of Contact:*

| Role | Name | Email | Responsibility |
|------|------|-------|----------------|
| Project Lead | [Name] | [email] | Overall accountability |
| Technical Lead | [Name] | [email] | Technical assessment |
| Business Lead | [Name] | [email] | Market/business analysis |
| Sponsor | [Name] | [email] | Executive oversight |

═══════════════════════════════════════════════════════════════════════════════
OUTPUT FORMAT
═══════════════════════════════════════════════════════════════════════════════

All monetary values in {currency}.

**Document Formatting Requirements:**
- Use clear section headers matching the structure above
- Tables should be properly formatted
- Bullet points for lists
- Bold for emphasis on key terms
- Professional, executive-ready tone

**IMPORTANT: Do NOT include a bibliography, references, or sources section at the end of your response. Source references should be inline within tables or text.**

═══════════════════════════════════════════════════════════════════════════════
PYTHON-DOCX SCRIPT GENERATION
═══════════════════════════════════════════════════════════════════════════════

After providing the report content above, generate a Python script using python-docx to create the Word document.

**CRITICAL RULES:**
• Document variable MUST be named exactly: doc
• Do NOT include doc.save() — the application handles file saving
• Use only python-docx library
• Match the formatting of the "Deep Dive Summary - Recommend to Explore" template

**Document Styling Requirements:**
• Title: "DeepDive Summary" — Large, bold, dark blue (RGB: 0, 103, 185)
• Section Headers: Bold, dark blue, larger font (14-16pt)
• Subsection Headers: Bold, slightly smaller (12pt)
• Body Text: Regular, 11pt
• Tables: Light blue header row (RGB: 74, 155, 217), alternating row colors
• Bullet points for lists
• Professional spacing between sections

**Available Imports:**
```python
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

Document Structure to Create:
Title: "DeepDive Summary"
Each section with proper header styling
All tables with formatted headers and data rows
Proper spacing and page breaks where appropriate
Header/footer with document metadata (optional)

Table Formatting:
Header row: Dark blue background (0, 103, 185), white text, bold
Data rows: Alternating white and light gray
Borders: Light gray, thin
Cell padding: Appropriate for readability 
"""


# =============================================================================
# COMPETITIVE ANALYSIS TEMPLATE
# =============================================================================

COMPETITIVE_ANALYSIS_TEMPLATE = """You are acting as a strategic business analyst specializing in competitive intelligence for the {industry} sector.

I need you to analyze the following innovation against the competitive landscape.

═══════════════════════════════════════════════════════════════════════════════
INNOVATION DETAILS
═══════════════════════════════════════════════════════════════════════════════

**Innovation Name:** {innovation_name}

**Innovation Description:** 
{innovation_description}

**Target Market/Industry:** {industry}

**Geographic Scope:** {geographic_scope}

═══════════════════════════════════════════════════════════════════════════════
ANALYSIS REQUIREMENTS
═══════════════════════════════════════════════════════════════════════════════

Please provide a comprehensive competitive analysis including:

1. **Market Landscape Overview**
   - Key players and market share
   - Market size and growth rate
   - Industry trends and drivers

2. **Direct Competitors** (Top 5)
   | Competitor | Product/Solution | Strengths | Weaknesses | Market Position |
   |------------|-----------------|-----------|------------|-----------------|

3. **Competitive Differentiation Matrix**
   | Feature | Our Innovation | Competitor A | Competitor B | Competitor C |
   |---------|---------------|--------------|--------------|--------------|

4. **SWOT Analysis**
   - Strengths
   - Weaknesses  
   - Opportunities
   - Threats

5. **Strategic Recommendations**
   - Market entry strategy
   - Positioning recommendations
   - Key success factors

All monetary values in {currency}.

**IMPORTANT: Do NOT include a bibliography, references, or sources section at the end of your response.**
"""


# =============================================================================
# PROMPT TEMPLATE CLASS
# =============================================================================

class PromptTemplate:
    """Represents a prompt template with metadata."""
    
    def __init__(
        self,
        id: str,
        name: str,
        description: str,
        version: str,
        author: str,
        created_date: str,
        category: str,
        required_fields: List[str],
        optional_fields: List[str],
        template_content: str
    ):
        self.id = id
        self.name = name
        self.description = description
        self.version = version
        self.author = author
        self.created_date = created_date
        self.category = category
        self.required_fields = required_fields
        self.optional_fields = optional_fields
        self.template_content = template_content
    
    def render(self, context: Dict[str, str]) -> str:
        """Render the template with the provided context."""
        try:
            return self.template_content.format(**context)
        except KeyError as e:
            raise ValueError(f"Missing required field in context: {e}")


# =============================================================================
# TEMPLATE REGISTRY CLASS
# =============================================================================

class TemplateRegistry:
    """
    Central registry for all prompt templates.
    """
    
    def __init__(self):
        self._templates: Dict[str, PromptTemplate] = {}
        self._load_default_templates()
    
    def _load_default_templates(self):
        """Load built-in default templates."""
        
        # Value Creation Model Template V3 (Enhanced with Growth Rates)
        self.register(PromptTemplate(
            id="value_creation_v3",
            name="Value Creation Model (Enhanced)",
            description="Comprehensive value quantification with 100% penetration, geographic breakdown, segmentation, monetization potential, growth rates, and 2026-2040 projections",
            version="3.0",
            author="Analog Garage",
            created_date="2024-01-01",
            category="Financial Analysis",
            required_fields=["innovation_name", "innovation_description", "industry"],
            optional_fields=[
                "geographic_scope", "analysis_timeframe", "innovation_stage", "currency",
                "problem_statement", "target_customer", "competitive_advantage",
                "tam", "target_penetration", "price_point", "regulatory_pathway",
                "ip_status", "key_risks", "key_assumptions"
            ],
            template_content=VALUE_CREATION_TEMPLATE_V3
        ))
        
        # Competitive Analysis Template
        self.register(PromptTemplate(
            id="competitive_analysis_v1",
            name="Competitive Analysis",
            description="Strategic competitive landscape analysis",
            version="1.0",
            author="Analog Garage",
            created_date="2024-01-01",
            category="Strategy",
            required_fields=["innovation_name", "innovation_description", "industry"],
            optional_fields=["geographic_scope", "currency"],
            template_content=COMPETITIVE_ANALYSIS_TEMPLATE
        ))
        
        # Deep Dive Summary - Recommend to Explore Template (NEW)
        self.register(PromptTemplate(
            id="gonogo_report_v1",
            name="Deep Dive Summary - Recommend to Explore",
            description="Comprehensive Deep Dive assessment document for Analog Garage exploration decisions",
            version="1.0",
            author="Analog Garage",
            created_date="2024-01-01",
            category="Analysis",
            required_fields=["innovation_name", "innovation_description", "industry"],
            optional_fields=[
                "geographic_scope", "analysis_timeframe", "innovation_stage", "currency",
                "problem_statement", "target_customer", "competitive_advantage",
                "tam", "target_penetration", "price_point", "regulatory_pathway",
                "ip_status", "key_risks", "key_assumptions"
            ],
            template_content=GONOGO_REPORT_TEMPLATE
        ))
    
    def register(self, template: PromptTemplate) -> None:
        """Register a new template."""
        self._templates[template.id] = template
    
    def get(self, template_id: str) -> Optional[PromptTemplate]:
        """Get a template by ID."""
        return self._templates.get(template_id)
    
    def get_all(self) -> List[PromptTemplate]:
        """Get all registered templates."""
        return list(self._templates.values())
    
    def get_by_category(self, category: str) -> List[PromptTemplate]:
        """Get templates filtered by category."""
        return [t for t in self._templates.values() if t.category == category]
    
    def list_ids(self) -> List[str]:
        """List all template IDs."""
        return list(self._templates.keys())
    
    def list_names(self) -> List[str]:
        """List all template display names."""
        return [t.name for t in self._templates.values()]
    
    def get_categories(self) -> List[str]:
        """Get unique categories."""
        return list(set(t.category for t in self._templates.values()))
    
    def unregister(self, template_id: str) -> bool:
        """Remove a template from the registry."""
        if template_id in self._templates:
            del self._templates[template_id]
            return True
        return False


# =============================================================================
# CREATE GLOBAL REGISTRY INSTANCE
# =============================================================================

template_registry = TemplateRegistry()


# =============================================================================
# ENHANCED PROMPT BUILDER
# =============================================================================

def build_enhanced_prompt(context: Dict[str, str]) -> str:
    """
    Build the enhanced V3 prompt with conditional sections based on available data.
    """
    
    # Build sections
    problem_section = ""
    if context.get("problem_statement"):
        problem_section = f"**Problem Statement:**\n{context['problem_statement']}"
    
    customer_section = ""
    if context.get("target_customer"):
        customer_section = f"**Target Customer:** {context['target_customer']}"
    
    advantage_section = ""
    if context.get("competitive_advantage"):
        advantage_section = f"**Competitive Advantage:** {context['competitive_advantage']}"
    
    # Market section
    market_items = []
    if context.get("tam"):
        market_items.append(f"• Total Addressable Market: {context['tam']}")
    if context.get("target_penetration"):
        market_items.append(f"• Target Market Penetration: {context['target_penetration']}")
    if context.get("price_point"):
        market_items.append(f"• Target Price Point: {context['price_point']}")
    
    market_section = "\n".join(market_items) if market_items else "(No market estimates provided — please derive from research)"
    
    # Regulatory section
    regulatory_items = []
    if context.get("regulatory_pathway"):
        regulatory_items.append(f"• Regulatory Pathway: {context['regulatory_pathway']}")
    if context.get("ip_status"):
        regulatory_items.append(f"• IP Status: {context['ip_status']}")
    
    regulatory_section = "\n".join(regulatory_items) if regulatory_items else "(No regulatory/IP context provided)"
    
    # Risks section
    risks_section = ""
    if context.get("key_risks"):
        risks_section = f"**Key Risks Identified:**\n{context['key_risks']}"
    else:
        risks_section = "(No specific risks identified — please assess based on industry and stage)"
    
    risks_assessment = ""
    if context.get("key_risks"):
        risks_assessment = f"Incorporate user-identified risks:\n{context['key_risks']}"
    else:
        risks_assessment = "Assess risks based on innovation stage, industry, and market factors."
    
    # Assumptions section
    assumptions_section = ""
    assumptions_note = ""
    if context.get("key_assumptions"):
        assumptions_section = f"**User-Provided Assumptions:**\n{context['key_assumptions']}"
        assumptions_note = f"Incorporate user-provided assumptions: {context['key_assumptions']}"
    else:
        assumptions_section = "(No specific assumptions provided)"
        assumptions_note = "Develop assumptions based on industry benchmarks and comparable innovations."
    
    # Geographic sections
    geographic_scope = context.get("geographic_scope", "Global")
    
    if geographic_scope.lower() == "global":
        geographic_breakdown_section = GEOGRAPHIC_BREAKDOWN_GLOBAL
        geographic_value_section = GEOGRAPHIC_VALUE_GLOBAL
    else:
        geographic_breakdown_section = GEOGRAPHIC_BREAKDOWN_REGIONAL.format(
            geographic_scope=geographic_scope
        )
        geographic_value_section = f"""
**6C. Regional Analysis ({geographic_scope}):**

Provide sub-regional breakdown as appropriate for the specified geography.
"""
    
    # Segmentation instructions
    industry = context.get("industry", "")
    has_segment_info = any([
        "segment" in industry.lower(),
        context.get("target_customer"),
        "segment" in context.get("innovation_description", "").lower()
    ])
    
    if has_segment_info:
        segmentation_instructions = SEGMENTATION_PROVIDED.format(industry=industry)
    else:
        segmentation_instructions = SEGMENTATION_NOT_PROVIDED
    
    # Customer note
    customer_note = ""
    if context.get("target_customer"):
        customer_note = f"\n*Note: User identified \"{context['target_customer']}\" as a key customer/stakeholder*"
    
    # Assemble final context
    final_context = {
        "innovation_name": context.get("innovation_name", ""),
        "innovation_description": context.get("innovation_description", ""),
        "industry": context.get("industry", ""),
        "geographic_scope": geographic_scope,
        "analysis_timeframe": context.get("analysis_timeframe", "Year 1 at Scale"),
        "innovation_stage": context.get("innovation_stage", "Concept"),
        "currency": context.get("currency", "USD"),
        "problem_section": problem_section,
        "customer_section": customer_section,
        "advantage_section": advantage_section,
        "market_section": market_section,
        "regulatory_section": regulatory_section,
        "risks_section": risks_section,
        "risks_assessment": risks_assessment,
        "assumptions_section": assumptions_section,
        "assumptions_note": assumptions_note,
        "geographic_breakdown_section": geographic_breakdown_section,
        "geographic_value_section": geographic_value_section,
        "segmentation_instructions": segmentation_instructions,
        "customer_note": customer_note,
    }
    
    return VALUE_CREATION_TEMPLATE_V3.format(**final_context)


# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

def get_template(template_id: str) -> Optional[PromptTemplate]:
    """Convenience function to get a template."""
    return template_registry.get(template_id)


def list_templates() -> List[Dict[str, str]]:
    """List all templates with basic info."""
    return [
        {
            "id": t.id,
            "name": t.name,
            "description": t.description,
            "category": t.category,
            "version": t.version,
        }
        for t in template_registry.get_all()
    ]


def render_template(template_id: str, context: Dict[str, str]) -> str:
    """Render a template with context."""
    template = template_registry.get(template_id)
    if not template:
        raise ValueError(f"Template not found: {template_id}")
    return template.render(context)
