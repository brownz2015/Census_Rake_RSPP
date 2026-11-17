# Custom Census Tract Demographics via Raking

An R tool for producing custom census tract–level demographic estimates that published American Community Survey (ACS) tables do not provide. The tool reweights ACS Public Use Microdata Sample (PUMS) person records to match each tract's published ACS marginal totals using iterative proportional fitting (raking). The resulting tract-specific weights can then be summed to produce cross-tabulations that the Census Bureau does not publish at the tract level, for example housing cost burden by language spoken.

Raking specifications are defined in three linked spreadsheet templates, so users can change which variables are raked and how they are binned without writing R code.

This tool accompanies:

> Brown, Z., Martens, N., & Bolter, K. (2026). Beyond Published Tables: Generating and Validating Custom Census Tract Demographics for Policy Analysis. *Regional Science Policy & Practice*. [ARTICLE DOI]

Replication data and scripts for the article's analyses are archived separately at Harvard Dataverse: [DATAVERSE DOI]

---

## Requirements

- R (developed and tested on R 4.x)
- A free Census API key: <https://api.census.gov/data/key_signup.html>
- R packages:

```r
install.packages(c("openxlsx", "stringr", "tidycensus", "qualV", "writexl",
                   "readxl", "dplyr", "varhandle", "tibble", "anesrake", "survey"))
```

## Setup

1. **Set your Census API key** once per machine:
   ```r
   tidycensus::census_api_key("YOUR_KEY", install = TRUE)
   readRenviron("~/.Renviron")
   ```
   The script reads the key from the `CENSUS_API_KEY` environment variable. Never commit a key to this repository.

2. **Set the working directory** to this repository's root folder. The script looks for the templates in `templates/`.

3. **Edit the configuration block** near the top of `Rake_Tracts.R`:

   | Setting | Meaning | Example |
   |---|---|---|
   | `aCode` | PUMA code(s) to process | `"02102"` |
   | `st` | State name | `"Michigan"` |
   | `cnty` | County or counties containing the target tracts | `"Kalamazoo"` |
   | `yr` | ACS data year | `2023` |
   | `microdatasurv` | PUMS sample | `"acs5"` |
   | `tract_list_by_puma` | Named list mapping each PUMA to the 11-digit GEOIDs of its tracts | see script |
   | `binsheetname`, `pumvarname`, `binvars` | Template file names (without `.xlsx`) | see below |

   The shipped configuration reproduces the article's Kalamazoo, Michigan application (PUMA 02102, 35 tracts).

4. **Run** `Rake_Tracts.R`.

## Output

`PUMA_Master` is a named list with one dataframe per PUMA. Each dataframe holds the PUMA's household-population PUMS person records plus one `scaled_weight_<GEOID>` column per tract. Each tract's weights are scaled to that tract's ACS total population.

To estimate a quantity for a tract, sum that tract's weight column over the records that meet your condition. For example:

```r
d <- PUMA_Master[["02102"]]
sum(d$scaled_weight_26077000100[d$AGEP >= 65], na.rm = TRUE)   # residents 65+ in tract 000100
```

To add PUMS variables that are not used in raking (e.g., `ENG`, `LANP`, `GRPIP`, `TEN`), pull them with `tidycensus::get_pums()` and join on `SERIALNO` and `SPORDER`.

## The spreadsheet templates

The interface consists of three linked files in `templates/` that together define a raking specification.

**Variable selection (`PUMA_Vars_FiveGroup.xlsx`).** A single-column template listing the raw PUMS variables to retrieve for a run. For the Kalamazoo application this list includes `AGEP`, `POVPIP`, `RAC1P`, `HISP`, `WKWN`, `WKHP`, `SEX`, and `SCHL`.

**Variable construction (`Conditionals_Final_FiveGroup.xlsx`).** One sheet per constructed variable. Each sheet specifies how raw PUMS values map to an internal bin code.
- *Categorical mappings* recode a single variable directly. For race, for example:

  | RAC1P | Bin |
  |---|---|
  | 1 | b1 |
  | 2 | b2 |
  | 3, 4, 5 | b3 |

- *Conditional and composite mappings* condition on age and sex and can combine several source columns. Employment status, for example, combines usual hours worked (`WKHP`) and weeks worked (`WKWN`):

  | AGEP min | AGEP max | SEX | WKHP range | WKWN range | Bin |
  |---|---|---|---|---|---|
  | 16 | 65 | 1 | 0–1 | 0–1 | d2 |
  | 16 | 65 | 1 | 1–35 | 1–53 | d3 |
  | 16 | 65 | 1 | 35–100 | 48–53 | d4 |

  Income-to-poverty ratio and educational attainment follow the same conditional structure. Their bins are restricted to the age range over which each ACS table's universe is defined.

**ACS table crosswalk (`Bin_Template_Final_FiveGroup.xlsx`).** Links each bin code to the ACS summary table cell it targets during raking:

  | Bin code | ACS cell |
  |---|---|
  | b1 (Race) | B02001_002 |
  | b2 (Race) | B02001_003 |
  | d2 (Employment) | B23022_025 |
  | d3 (Employment) | B23022_011 |

Age-restricted variables, such as income-to-poverty ratio and educational attainment, place individuals outside the relevant age range in a separate out-of-universe bin rather than in a substantive category. This keeps them aligned with the universe definitions of the corresponding ACS tables.

**Note:** `Rake_Tracts.R` currently reads exactly five construction sheets in a fixed order (race/ethnicity, income-to-poverty, education, employment, age-sex), and applies the out-of-universe bins for those variables in code. Adding or reordering raking dimensions requires matching changes to the script.

## Raking settings

Raking uses `anesrake` with `cap = 10`, `convcrit = 0.01`, `type = "pctlim"`, `pctlim = 0.05`, and `choosemethod = "total"`.

## Scope and limitations

- **Household population only.** Group-quarters records (`TYPEHUGQ != 1`) are dropped before raking. Estimates therefore do not represent people in dormitories, correctional facilities, nursing homes, or military barracks. In the article's validation, every tract that failed to converge either housed its entire population in group quarters or had a total population under 25.
- Estimates inherit the sampling error of both the ACS tables and the PUMS. In the article's uncertainty analysis, subpopulations of roughly 20–80 people per tract fell mostly in the caution range (CV 0.15–0.30). Cells under 20 people were mostly unreliable and are best aggregated across tracts.
- Accuracy depends on the PUMA's microdata containing people who resemble each tract's residents. Characteristics that are not correlated with the raking variables will not be corrected by reweighting.

## Citation

If you use this tool, please cite the article above.

## License

[LICENSE: choose one, e.g. MIT, and add a LICENSE file]
