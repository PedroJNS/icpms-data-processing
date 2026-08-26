# ICP-MS Data Processing Tool (Agilent 7900)

Python tool for automated data processing, concentration calculations (converting raw readings from ppb to ppm and %), and analytical data treatment for **Agilent 7900 ICP-MS (MassHunter)** outputs.

## 📌 Ownership & Licensing
* **Author:** Pedro J. (PedroJNS)
* **License:** GNU General Public License v3.0 (GPL-3.0)
* **Development:** Personal and independent open-source project.

## 🚀 Key Features
* Vectorized data processing using `pandas` and `numpy`.
* Automated conversion from solution concentrations ($\mu\text{g/L}$ or ppb) to solid sample concentrations ($\text{mg/kg}$ or ppm, and $\%$).
* Flexible CLI workflow for batch sample digestion input (Sample Mass in mg and Dilution Volume in mL).

## 🧮 Applied Formulas
* **ppm ($\text{mg/kg}$):** `(ICP_Reading_ppb * Dilution_Volume_mL) / Sample_Mass_mg`
* **Percentage ($\%$):** `ppm / 10000`

## ⚠️ Disclaimer
This code is provided "as is", without warranty of any kind. The author assumes no liability for calculation errors resulting from improperly formatted input data or misuse in laboratory analytical procedures.
