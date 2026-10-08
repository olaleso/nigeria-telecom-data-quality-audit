# Nigeria Telecom Data Quality Audit

Reproducibility package for the study:

**Auditing Public Telecommunications Data Quality in Nigeria: A Reproducible ISO/IEC 25012 Assessment and DAMA-DMBOK-Informed Governance Framework**

**Authors:** Sowunmi Ibrahim Olaleye; Yusuff Babatunde Olujobi  
**Affiliation:** Department of Computer Science, Caleb University, Lagos, Nigeria  
**Target journal:** *Data* (MDPI)

## Overview

This repository contains the final V3 Python audit pipeline and reproducibility materials used to assess public Nigerian telecommunications workbooks using selected ISO/IEC 25012 data-quality characteristics, with a DAMA-DMBOK-informed governance interpretation and the proposed Telecom Regulatory Data Quality Governance Framework (TR-DQGF).

The study analyses seven National Bureau of Statistics (NBS) telecommunications workbooks covering **Q2 2024 to Q4 2025**.

## Audit coverage reported in the manuscript

- 7 workbooks
- 84 worksheet occurrences
- 21 canonical sheets
- 518 canonical Voice/Internet state/FCT records
- 1,036 YoY/QoQ calculations
- 666 cross-period reference checks
- 14 canonical Porting reconciliations
- 84 geopolitical-zone reconciliations
- 444 historical state-version comparisons

## Principal reported results

- Normalised state/FCT completeness: **100.00%**
- State subtotal reconciliation: **100.00%**
- YoY/QoQ recalculation: **100.00%**
- Cross-period consistency: **99.85% (665/666)**
- Porting component-to-total reconciliation: **100.00% (14/14)**
- Zone aggregate reconciliation: **84.52% (71/84)**
- Historical state-total stability: **100% (444/444)**
- Canonical sheets with embedded source/note markers: **0/21**

The principal aggregate anomaly is a reciprocal pattern between the **Q3 2025 Voice** and **Q4 2025 Internet** geopolitical-zone series.

## Repository structure

```text
.
├── README.md
├── LICENSE
├── CITATION.cff
├── .zenodo.json
├── requirements.txt
├── code/
│   └── run_nbs_iso25012_audit_v3.py
├── metadata/
│   └── v3_audit_metadata.json
└── outputs/
    └── README.md
```

## Software environment

The manuscript reports:

- Python **3.12.10**
- openpyxl **3.1.5**

Install the dependency with:

```bash
python -m pip install -r requirements.txt
```

## Source data

The original source workbooks are third-party public data from the National Bureau of Statistics (Nigeria), **Nigerian Telecoms Data**, reference ID **NGA-NBS-TELECOMS**.

Catalogue record:

https://microdata.nigerianstat.gov.ng/index.php/catalog/179

This repository does not assert ownership of the original NBS source workbooks. Where redistribution is restricted or unnecessary, obtain the source files directly from the authoritative catalogue.

## Running the V3 audit

The archived V3 script preserves the study environment path used during the analysis:

```python
ROOT = Path(r"C:\PhDDataSets\NBS")
```

The script expects the seven extracted Excel workbooks below:

```text
C:\PhDDataSets\NBS\extracted
```

and writes outputs to:

```text
C:\PhDDataSets\NBS\iso25012_audit_v3
```

For another computer, edit the `ROOT` path before execution.

Run:

```bash
python code/run_nbs_iso25012_audit_v3.py
```

The V3 pipeline contains a fail-safe and terminates if all seven expected workbooks are not parsed.

## Generated outputs

The V3 script generates:

1. `v3_01_sheet_occurrences.csv`
2. `v3_02_state_sheet_metrics.csv`
3. `v3_03_canonical_state_data.csv`
4. `v3_04_cross_period_reconciliation.csv`
5. `v3_05_porting_reconciliation.csv`
6. `v3_06_zone_reconciliation.csv`
7. `v3_07_revision_stability.csv`
8. `v3_08_formula_transparency_drift.csv`
9. `v3_09_schema_evolution.csv`
10. `v3_10_metadata_traceability.csv`
11. `v3_11_issue_register.csv`
12. `v3_12_dimension_summary.csv`
13. `v3_audit_metadata.json`

The final machine-readable CSV outputs should be added to `outputs/` before the archival Zenodo release is published.

## Methodological safeguards

- Internal reconciliation is not treated as proof of independent real-world accuracy.
- Blank cells, hyphens and zeros are distinguished.
- `NASSARAWA` is normalised to `NASARAWA` for joins and aggregation while retained as a reference-data issue.
- Currentness is not inferred from local file timestamps.
- Compliance is not scored without an explicit external NCC/NBS benchmark.
- Data-quality measures are reported separately rather than collapsed into a composite score.
- Historical workbook copies are retained for version-stability and formula-transparency analysis.

## Zenodo archival workflow

This GitHub repository is intended to be archived through Zenodo after the final manuscript reproducibility package has been completed.

1. Add the final V3 machine-readable outputs to `outputs/`.
2. Create a GitHub release, recommended tag: **v1.0.0**.
3. Archive that release in Zenodo.
4. Mint the Zenodo DOI.
5. Insert the DOI into the manuscript's Data Availability and Code Availability statements.

## Citation

Citation metadata are provided in `CITATION.cff`. The Zenodo DOI should be added after it has been minted.

## Licence

The software code is released under the MIT License. Third-party source data remain subject to the terms of the original data provider.
