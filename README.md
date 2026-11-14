# Fortius FloodRapid: Parametric Flood Pricing for Under-Served Households

## Repository Contents

| File | Description |
| :--- | :--- |
| [`report/Fortius_FloodRapid_Pricing_Report.pdf`](./report/Fortius_FloodRapid_Pricing_Report.pdf) | **Executive & Technical Report (9 Pages):** Complete mathematical explanation of the Version 10 stochastic model and its limitations, pricing and payout sensitivity analyses, optimal product design recommendation (£1,300 premium / £20,000 payout), and gauge basis risk stress testing. |
| [`model/FortiusFloodRapidV10_Solution.xlsx`](./model/FortiusFloodRapidV10_Solution.xlsx) | **Stochastic Pricing & Reserving Workbook (12 Sheets, Zero VBA):** The 7 core model sheets (`Overview`, `Inputs & Parameters`, `Uniform Random Numbers`, `Simulations`, `Cashflows & Reserves`, `Simulations Summary`, `Pricing Dashboard`) extended with 5 analytical sheets (`Default Distribution Histograms`, `Premium Sensitivity Testing`, `Pricing Strategy`, `Preferred Combination`, `Sensitivity Testing Basis Risk`). |

---

## Executive Context & Problem Statement

**Fortius Insurance** is a UK household flood insurer evaluating its highest-risk segment: **High-risk flood zone ("Zone H")**, comprising **1,000 households spread across 10 river valleys** ($V_{\text{valleys}} = 10$), where each household faces a **1-in-25 year ($T = 25$, or $4\%$ annual)** probability of flooding.

The existing product, **"Fortius Flood"**, is a traditional indemnity policy with an average payout of **$K_{\text{std}} = £30,000$** (per-claim log-normal severity volatility $\sigma = 0.08$), fixed admin cost of **$£50/\text{policy/yr}$**, manual claim-handling cost of **$£1,500/\text{claim}$**, and a standard risk-priced annual premium of **$P_{\text{std}} = £2,640$**. Take-up is heavily skewed toward high-income households, leaving lower-income households largely priced out of protection:

| Household Income Bucket ($i$) | Zone H Population ($N_i$) | Annual Income ($I_i$) | Household Wealth ($W_i$) | Traditional-Only Take-Up Rate ($w_{i,\text{std}}$) | Traditional-Only Policies (`TRAD_BOOK`) |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Low Income** | $500$ ($50\%$) | $£15,000$ | $£30,000$ | $0.81\%$ ($0.8\%$) | $4.05$ |
| **Middle Income** | $400$ ($40\%$) | $£40,000$ | $£100,000$ | $30.78\%$ ($30.8\%$) | $123.10$ |
| **High Income** | $100$ ($10\%$) | $£80,000$ | $£200,000$ | $46.43\%$ ($46.4\%$) | $46.43$ |
| **Total / Portfolio** | **$1,000$** | — | — | **$17.36\%$** | **$173.58$** |

To extend coverage to under-served Low- and Middle-income households while maintaining a sensible **Return on Capital ($\text{RoC}$)**, Fortius is evaluating **"Fortius Flood Rapid"** — an optional **parametric flood insurance product** paying a fixed lump sum (**$K_{\text{par}} = £20,000$** default per trigger) via an automated household flood gauge, with **no manual claim handling ($£0/\text{claim}$)**, lower fixed admin overhead (**$£10/\text{policy/yr}$**), and a default premium rate of **$P_{\text{par}} = £1,700/\text{yr}$**.

---

## Mathematical & Stochastic Model Formulation

All equations below correspond directly to the formulas implemented across the workbook sheets (`Inputs & Parameters`, `Simulations`, `Cashflows & Reserves`, `Simulations Summary`, and `Pricing Dashboard`) and Section 2 of the report.

### 1. Customer Demand & Willingness-to-Buy Model (`Inputs & Parameters`, rows 27–72)

For each income bucket $i \in \{\text{High}, \text{Middle}, \text{Low}\}$ and product $k \in \{\text{std}, \text{par}\}$ with annual premium $P_k$ and payout per claim/trigger $K_k$, a household's willingness-to-buy score $w_{i,k} \in [0, 1]$ is built in three steps:

1. **Affordability Score $a_{i,k} \in [0, 1]$ (`cells D56:E58`):** A logistic function of the premium-to-income ratio $P_k / I_i$ centred at affordability threshold $\theta_a = 0.07$ ($7\%$ of income) with steepness $s_a = 0.025$:
   $$a_{i,k} = \frac{1}{1 + \exp\!\left(\dfrac{P_k / I_i - \theta_a}{s_a}\right)}$$

2. **Value Score $v_{i,k} \in [0, 1]$ (`cells F56:G58`):** Combines a penalty for excess premium loading above the pure risk cost $K_k / T$ (adjusted for risk aversion $\eta = 2$, household wealth $W_i$, and value tolerance $V = 3$) with a coverage adequacy factor relative to reference full-damage loss $L = K_{\text{std}} = £30,000$ and reference wealth $W_{\text{ref}} = £100,000$:
   $$v_{i,k} = \max\!\left(0,\; 1 - \frac{1}{V}\max\!\left(0,\; \frac{P_k}{K_k / T} - 1 - \frac{\eta}{W_i}\left(1 - \frac{1}{T}\right)\right)\right) \times \min\!\left(1,\; \left(\frac{K_k}{L}\right)^{W_i / W_{\text{ref}}}\right)$$

3. **Willingness-to-Buy Score $w_{i,k} \in [0, 1]$ (`cells H56:I58`):** Scaled by Zone H insurance appetite (`Zone taste`) $\tau = 0.95$:
   $$w_{i,k} = \min\!\big(1,\; \tau \cdot a_{i,k} \cdot v_{i,k}\big)$$

4. **Portfolio Take-Up under Both Scenarios (`cells C62:G72`):**
   - **Traditional-Only Scenario (`TRAD_BOOK`):** Only the standard indemnity product is offered:
     $$\text{Book}_{i,\text{trad}} = N_i \cdot w_{i,\text{std}}$$
   - **Traditional + Parametric Scenario (`STD_BOOK` & `PAR_BOOK`):** Households compare both products. The probability of buying *any* flood cover is:
     $$P_i = 1 - \big(1 - w_{i,\text{std}}\big)\big(1 - w_{i,\text{par}}\big) \tag{1}$$
     and policies split between the Standard Indemnity book and the Parametric Rapid book in proportion to relative willingness-to-buy:
     $$\text{Book}_{i,\text{std}} = N_i \cdot P_i \cdot \frac{w_{i,\text{std}}}{w_{i,\text{std}} + w_{i,\text{par}}}, \qquad \text{Book}_{i,\text{par}} = N_i \cdot P_i \cdot \frac{w_{i,\text{par}}}{w_{i,\text{std}} + w_{i,\text{par}}} \tag{2}$$

---

### 2. Spatially Correlated Valley Flood Index & Two-Stage Flood Threshold (`Simulations`, col C)

1. **Latent Valley Water Index $W_{v,t}$:** Across $1,000$ Monte Carlo simulations and $t \in \{1,\dots,5\}$ projection years, each valley $v \in \{1,\dots,10\}$ experiences a latent water level driven by a UK-wide weather factor $Z_{\text{UK},t} = \Phi^{-1}(U_{\text{UK},t})$, a valley residual $Z_{v,t} = \Phi^{-1}(U_{v,t})$, valley correlation $\rho = 0.30$, and linear annual climate drift $d = 0.008$:
   $$W_{v,t} = \rho Z_{\text{UK},t} + \sqrt{1 - \rho^2}\,Z_{v,t} + (t - 1)d, \qquad Z_{\text{UK},t}, Z_{v,t} \stackrel{\text{i.i.d.}}{\sim} \mathcal{N}(0, 1)$$
   *(Note: Report Section 2.4.1 writes the residual weight as $\sqrt{1-\rho}\,Z_{v,t}$ per the audit trail notation; `Simulations!C9` implements `SQRT(1-RHO^2)` so that $\text{Var}(W_{v,1}) = 1$.)*

2. **Two-Stage Flood Threshold (`THR`, `Inputs & Parameters!B21`):** Conditional on a valley flooding, each household inside that valley floods with within-event probability $p_{\text{within}} = 0.80$. To match the 1-in-25 year ($1/T = 4\%$) unconditional household flood risk in Year 1, each valley's Year-1 flood probability is $\frac{1}{T \cdot p_{\text{within}}} = \frac{1}{25 \times 0.80} = 0.05$ ($5\%$), giving threshold:
   $$\text{THR} = \Phi^{-1}\!\left(1 - \frac{1}{T \cdot p_{\text{within}}}\right) = \Phi^{-1}(0.95) \approx 1.6449, \qquad F_t = \sum_{v=1}^{10} \mathbb{I}\big(W_{v,t} > \text{THR}\big)$$
   where $F_t \in \{0,\dots,10\}$ is the number of flooded valleys in year $t$.

---

### 3. Within-Event Claim Counts & Parametric Gauge Basis Risk (`Simulations`, cols D–I)

Assuming policies are distributed evenly across the $V_{\text{valleys}} = 10$ valleys, $F_t$ flooded valleys expose $\text{round}\!\big(F_t \cdot \frac{\text{Book}_b}{10}\big)$ policies in book $b \in \{\text{trad}, \text{std}, \text{par}\}$.

1. **Actual Flooded Policies (`cols D–F`):** Drawn via inverse Binomial using a shared within-event uniform draw $U_{\text{within},t} \sim \text{Uniform}(0,1)$:
   $$N_{b,t}^{\text{flood}} = F_{\text{Binom}}^{-1}\!\left(U_{\text{within},t};\; \text{round}\!\left(\frac{F_t \cdot \text{Book}_b}{10}\right),\; p_{\text{within}}\right)$$
   where $N_{\text{trad},t}^{\text{flood}}$ (`col D`) and $N_{\text{std},t}^{\text{flood}}$ (`col E`) are the traditional indemnity claim counts under the two scenarios, and $N_{\text{par},t}^{\text{flood}}$ (`col F`) is the count of genuinely flooded parametric households.

2. **Parametric Gauge Triggers & Basis Risk Counts (`cols G–I`):**
   - **True Positives ($\text{TP}_t$, `col G`):** Flooded parametric households whose gauge rightly triggers with probability $\text{TPR} = 0.995$ (Gauge Sensitivity):
     $$\text{TP}_t = F_{\text{Binom}}^{-1}\!\big(U_{\text{TP},t};\; N_{\text{par},t}^{\text{flood}},\; \text{TPR}\big)$$
   - **False Positives ($\text{FP}_t$, `col H`):** Actually dry parametric households across *all* of Zone H ($\max(0, \text{round}(\text{Book}_{\text{par}}) - N_{\text{par},t}^{\text{flood}})$) whose gauge spuriously triggers with probability $1 - \text{TNR} = 1 - 0.998 = 0.002$ (complement of Gauge Specificity):
     $$\text{FP}_t = F_{\text{Binom}}^{-1}\!\big(U_{\text{FP},t};\; \max\!\big(0, \text{round}(\text{Book}_{\text{par}}) - N_{\text{par},t}^{\text{flood}}\big),\; 1 - \text{TNR}\big)$$
   - **False Negatives ($\text{FN}_t$, `col I`):** Flooded parametric households whose gauge fails to trigger:
     $$\text{FN}_t = N_{\text{par},t}^{\text{flood}} - \text{TP}_t$$

---

### 4. Aggregate Severity, 5-Lag Dirichlet Development & Reserving (`Simulations` cols J–W & `Cashflows & Reserves`)

1. **Traditional Indemnity Ultimate Loss (`cols J–K`):** Each traditional claim has mean payout $K_{\text{std}} = £30,000$ multiplied by a log-normal severity factor $S_j$ with mean $1$ and standard deviation $\sigma = 0.08$. For $n = N_{b,t}^{\text{flood}}$ claims, aggregate ultimate loss is approximated via the Central Limit Theorem (sharing uniform draw $U_{\text{sev},t}$ across both traditional books):
   $$\text{Ult}_{b,t} = K_{\text{std}}\sum_{j=1}^{n} S_j \approx K_{\text{std}}\Big(n + \sigma\sqrt{n}\,\Phi^{-1}(U_{\text{sev},t})\Big), \qquad b \in \{\text{trad}, \text{std}\}$$

2. **Parametric Instant Payout (`col L`):** Paid immediately in the trigger year with zero reserve and zero claim-handling cost:
   $$\text{Cash}_{\text{par},t} = \text{Incurred}_{\text{par},t} = \big(\text{TP}_t + \text{FP}_t\big) \cdot K_{\text{par}}$$

3. **5-Lag Dirichlet Claims Development (`Inputs & Parameters` rows 81–87 & `Simulations` cols M–W):** Traditional claims settle across 5 development lags $\ell \in \{0, 1, 2, 3, 4\}$ (Origin Year lag 0 plus 4 development years) with Dirichlet weights $\boldsymbol{\alpha} = (\alpha_0, \alpha_1, \alpha_2, \alpha_3, \alpha_4) = (10, 1, 0.5, 0.25, 0.1)$ constructed from independent $\text{Gamma}(\alpha_\ell, 1)$ draws:
   $$G_{t,\ell} = F_{\text{Gamma}}^{-1}\big(U_{t,\ell};\; \alpha_\ell, 1\big) \sim \text{Gamma}(\alpha_\ell, 1), \qquad d_{t,\ell} = \frac{G_{t,\ell}}{\sum_{m=0}^{4} G_{t,m}}, \qquad \bar{d}_\ell = \frac{\alpha_\ell}{\sum_{m=0}^{4}\alpha_m}$$
   giving expected payment fractions $\bar{\mathbf{d}} = (84.39\%, 8.44\%, 4.22\%, 2.11\%, 0.84\%)$ and tail reserve fractions $\bar{\tau}_L = \sum_{\ell=L+1}^{4} \bar{d}_\ell$ (`Cashflows & Reserves!C15:G15`).

4. **Pre-Launch Warm-Up Run-Off, Reserves & Incurred Loss (`Cashflows & Reserves`):** To reflect pre-existing traditional business entering Year 1, a deterministic steady-state warm-up stream based on $\lambda = \frac{\text{Book}_{\text{trad}}}{T} K_{\text{std}} = £208,297/\text{yr}$ sets the opening reserve $R_0 = \lambda \sum_{\ell=1}^{4} \ell\,\bar{d}_\ell = £55,370$, warm-up cash $\text{WarmCash}_t = \lambda \sum_{\ell=t}^{4}\bar{d}_\ell$, and warm-up reserve $\text{WarmRes}_t = \lambda \sum_{\ell=t+1}^{4}(\ell - t)\bar{d}_\ell$. In payment year $t \in \{1,\dots,5\}$:
   $$\text{Cash}_{b,t} = \text{WarmCash}_t + \sum_{s=1}^{t} \text{Ult}_{b,s}\,d_{s,\,t-s}, \qquad R_{b,t} = \text{WarmRes}_t + \sum_{s=1}^{t} \text{Ult}_{b,s}\,\bar{\tau}_{t-s}$$
   $$\text{Incurred}_{b,t} = \text{Cash}_{b,t} + R_{b,t} - R_{b,t-1}$$

---

### 5. Technical Profit, Required Capital ($99.5\%$ VaR) & Return on Capital (`Pricing Dashboard`)

Over the 5-year horizon, total per-simulation incurred loss is $L_{5\text{yr}}^{\text{inc}} = \sum_{t=1}^{5}\big(\text{Incurred}_{\text{std},t} + \text{Incurred}_{\text{par},t}\big)$ (`Simulations Summary` cols B & E). Expected 5-year Technical Profit, Required Risk Capital proxy ($99.5\%$ VaR), and Annualised Return on Capital ($\text{RoC}$) are defined as:

$$\text{TP}_{5\text{yr}} = 5\big(\text{Book}_{\text{std}}P_{\text{std}} + \text{Book}_{\text{par}}P_{\text{par}}\big) - \mathbb{E}\big[L_{5\text{yr}}^{\text{inc}}\big] - 5\big(\text{Book}_{\text{std}}c_{\text{admin,std}} + \text{Book}_{\text{par}}c_{\text{admin,par}}\big) - 5\,\mathbb{E}[N_{\text{std},t}]\,c_{\text{handle,std}}$$

$$\text{Capital} = \text{VaR}_{0.995}\big(L_{5\text{yr}}^{\text{inc}}\big), \qquad \text{RoC} = \frac{\text{TP}_{5\text{yr}} / 5}{\text{VaR}_{0.995}\big(L_{5\text{yr}}^{\text{inc}}\big)}$$

---

## Principal Model Limitations

1. **Monte Carlo Sampling Noise in the $99.5\%$ VaR Tail:** With $1,000$ simulations, the $99.5\%$ VaR is determined by the 5th–6th worst realisations, carrying $\approx 5\%$ Monte Carlo estimation noise (inflated by the single-zone concentration). Headline capital and $\text{RoC}$ figures are reliable to at most 2 significant figures.
2. **Third-Party Demand Model Extrapolation:** Demand parameters ($\theta_a, s_a, V, \eta, W_{\text{ref}}, \tau$) are calibrated externally on private commercial datasets and held fixed, making take-up predictions less reliable at extreme premium or payout combinations.
3. **Homogeneous Spatial Correlation ($\rho = 0.30$):** Every pair of valleys shares an identical equi-correlation $\rho = 0.30$ via the single UK factor $Z_{\text{UK},t}$, ignoring spatial adjacency where neighbouring river valleys flood together more frequently than distant ones.
4. **Deterministic Linear Climate Drift ($d = 0.008$):** Adding $(t-1)d$ linearly to $W_{v,t}$ induces a deterministic, accelerating increase in tail exceedance probability $\mathbb{P}(W_{v,t} > \text{THR})$ without parameter uncertainty.
5. **Shared Single Severity Multiplier & Independent Dirichlet Draws:** A single normal draw $U_{\text{sev},t}$ scales severity across flooded claims within a year, while Dirichlet development vectors $(d_{t,0},\dots,d_{t,4})$ are drawn independently each origin year, omitting multi-year post-catastrophe repair inflation and supply-chain bottlenecks.
6. **Uncorrelated Gauge Failures:** False positives and false negatives are drawn as independent Binomial trials across households each year, omitting systematic batch defects, shared power/telecom outages, or debris fouling during storms.

---

## Empirical Results, Pricing Optimization & Basis Risk Analysis

### 1. Premium Sensitivity Analysis (`Premium Sensitivity Testing`)

Holding parametric payout at $K_{\text{par}} = £20,000$, $\text{TPR} = 0.995$, and $\text{TNR} = 0.998$, varying the parametric annual premium $P_{\text{par}}$ from $£1,700$ down to $£500$ under the **Traditional + Parametric** scenario yields:

| Parametric Premium ($P_{\text{par}}$) | Low Income Par Take-Up | Middle Income Par Take-Up | High Income Par Take-Up | Low + Middle Par Policies | 5-Year Technical Profit | $99.5\%$ VaR (Capital) | Annual $\text{RoC}$ |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **£1,700 (Default)** | $39.18$ ($39$) | $100.84$ ($101$) | $19.54$ ($20$) | $140.02$ ($140$) | $£1,567,536$ ($£1.57\text{m}$) | $£4,991,863$ ($£4.99\text{m}$) | **$6.28\%$ ($6.3\%$)** |
| **£1,400** | $88.35$ ($88$) | $126.47$ ($126$) | $23.26$ ($23$) | $214.82$ ($215$) | $£1,531,067$ ($£1.53\text{m}$) | $£5,919,502$ ($£5.92\text{m}$) | **$5.17\%$ ($5.2\%$)** |
| **£1,100** | $170.41$ ($170$) | $152.61$ ($153$) | $26.93$ ($27$) | $323.02$ ($323$) | $£1,276,901$ ($£1.28\text{m}$) | $£7,245,044$ ($£7.25\text{m}$) | **$3.52\%$ ($3.5\%$)** |
| **£800** | $275.69$ ($276$) | $178.88$ ($179$) | $30.54$ ($31$) | $454.57$ ($455$) | $£653,450$ ($£0.65\text{m}$) | $£8,985,244$ ($£8.99\text{m}$) | **$1.45\%$ ($1.5\%$)** |
| **£500** | $339.02$ ($339$) | $184.08$ ($184$) | $30.84$ ($31$) | $523.10$ ($523$) | $-£193,281$ ($-£0.19\text{m}$) | $£9,845,444$ ($£9.85\text{m}$) | **$-0.39\%$ ($-0.4\%$)** |

Low-income households exhibit the steepest price elasticity (expanding nearly $9\times$ from $39$ to $339$ policies as premium drops from $£1,700$ to $£500$), whereas varying the payout trigger $K_{\text{par}}$ from $£14,000$ to $£30,000$ at $£1,700$ premium (`Pricing Strategy!A10:E14`) primarily attracts middle-income households because low-income demand is bottlenecked by the affordability score $a_{\text{Low},\text{par}}$.

---

### 2. Optimal Product Design Recommendation (`Pricing Strategy`)

Comparing candidate $(P_{\text{par}}, K_{\text{par}})$ combinations against the Traditional-Only baseline ($173.58$ total policies, $£1.11\text{m}$ 5-yr technical profit, $£3.37\text{m}$ $99.5\%$ VaR, $6.57\%$ RoC):

| Premium ($P_{\text{par}}$) | Payout ($K_{\text{par}}$) | Low + Middle Par Take-Up | High Par Take-Up | 5-Year Technical Profit | $99.5\%$ VaR (Capital) | Annual $\text{RoC}$ | Strategic Assessment |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **£1,700** | **£25,000** | $200.08$ ($200$) | $34.24$ ($34$) | $£1,595,679$ ($£1.60\text{m}$) | $£6,494,744$ ($£6.5\text{m}$) | $4.91\%$ ($4.9\%$) | Higher cover attracts Middle/High income; low Low-income reach ($51$ policies) |
| **£1,700** | **£20,000** | $140.02$ ($140$) | $19.54$ ($20$) | $£1,567,536$ ($£1.57\text{m}$) | $£4,991,863$ ($£5.0\text{m}$) | $6.28\%$ ($6.3\%$) | Default parameters; high $\text{RoC}$ but weak Low-income inclusion ($39$ policies) |
| **£1,100** | **£20,000** | $323.02$ ($323$) | $26.93$ ($27$) | $£1,276,901$ ($£1.28\text{m}$) | $£7,245,044$ ($£7.2\text{m}$) | $3.52\%$ ($3.5\%$) | Strong inclusion ($170$ Low, $153$ Mid), but $\text{RoC}$ diluted to $3.5\%$ |
| **£1,300 (Recommended)** | **£20,000** | **$247.23$ ($247$)** | **$24.49$ ($24$)** | **$£1,480,352$ ($£1.48\text{m}$)** | **$£6,299,702$ ($£6.3\text{m}$)** | **$4.70\%$ ($4.7\%$)** | **Optimal balance: $112$ Low + $135$ Mid parametric policies ($415$ total book) at $4.7\%$ $\text{RoC}$** |

---

### 3. Basis Risk Stress Testing (`Sensitivity Testing Basis Risk`)

Evaluated at the recommended **$(P_{\text{par}} = £1,300,\; K_{\text{par}} = £20,000)$** design, stressing Gauge Sensitivity ($\text{TPR}$) and Gauge Specificity ($\text{TNR}$) in $-5\%$ increments reveals a stark asymmetry between customer/regulatory risk and insurer solvency risk:

#### Stressing Gauge Sensitivity ($\text{TPR}$) — False Negatives ($\text{TNR} = 0.998$ fixed)

| Stress | Gauge Sensitivity ($\text{TPR}$) | Mean False Negatives (5-yr) | 99th Percentile FN | Max FN | Annual $\text{RoC}$ |
| :---: | :---: | :---: | :---: | :---: | :---: |
| **$0\%$ (Default)** | **$0.995$** | $0.29$ | $2$ | $4$ | **$4.7\%$** |
| **$-5\%$** | **$0.945$** | $3.15$ | $11$ | $15$ | **$5.0\%$** |
| **$-10\%$** | **$0.896$** | $5.93$ | $19$ | $23$ | **$5.3\%$** |
| **$-15\%$** | **$0.846$** | $8.78$ | $27$ | $32$ | **$5.7\%$** |
| **$-20\%$** | **$0.796$** | $11.63$ | $34$ | $40$ | **$6.1\%$** |
| **$-25\%$** | **$0.746$** | $14.47$ | $41$ | $50$ | **$6.4\%$** |

Worsening $\text{TPR}$ perversely *increases* insurer $\text{RoC}$ (from $4.7\%$ to $6.4\%$) because Fortius collects premiums while failing to pay genuinely flooded households—creating severe reputational and regulatory failure.

#### Stressing Gauge Specificity ($\text{TNR}$) — False Positives ($\text{TPR} = 0.995$ fixed)

| Stress | Gauge Specificity ($\text{TNR}$) | Mean False Positives (5-yr) | 99th Percentile FP | Max FP | Annual $\text{RoC}$ |
| :---: | :---: | :---: | :---: | :---: | :---: |
| **$0\%$ (Default)** | **$0.998$** | $2.62$ | $7$ | $9$ | **$4.7\%$** |
| **$-5\%$** | **$0.948$** | $67.63$ | $88$ | $94$ | **$0.5\%$** |
| **$-10\%$** | **$0.899$** | $131.38$ | $159$ | $168$ | **$-2.5\%$** |
| **$-15\%$** | **$0.849$** | $196.53$ | $230$ | $241$ | **$-4.9\%$** |
| **$-20\%$** | **$0.799$** | $261.65$ | $300$ | $314$ | **$-6.7\%$** |

Because every non-flooded parametric policy across all 10 valleys is exposed to a $1 - \text{TNR}$ false-trigger draw every year, even a $5\%$ drop in $\text{TNR}$ (from $0.998$ to $0.948$) multiplies mean 5-year false positives $26\times$ (from $2.62$ to $67.63$) and wipes out return on capital ($4.7\% \to 0.5\%$). At $\text{TNR} = 0.799$, $\text{RoC}$ plunges to **$-6.7\%$**—making high gauge specificity the single most critical financial dependency of the product.
