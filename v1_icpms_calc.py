"""
==============================================================================
ICP-MS Data Processing Module (Agilent 7900)
Author: Pedro J. (PedroJNS)
License: GNU General Public License v3.0 (GPL-3.0)
Description: Automated processing and conversion of raw ICP-MS solution 
             readings (ppb) to solid sample concentrations (ppm and %).
==============================================================================
"""

import pandas as pd
import numpy as np


def get_digestion_data_and_calculate_concentrations(df_clean_samples, digestion_data=None):
    """
    Calculates real concentrations in solid samples from ICP-MS solution readings (ppb).

    Unit conversions and formulas:
    - 1 ppb = 1 µg/L
    - Concentration (ppm or mg/kg) = [ICP Reading (ppb) * Dilution Volume (mL)] / Sample Mass (mg)
    - Concentration (%) = Concentration (ppm) / 10,000

    Parameters:
    -----------
    df_clean_samples : pd.DataFrame
        Cleaned raw DataFrame from Agilent 7900 MassHunter containing metadata in 
        the first two rows and concentration data (ppb) starting from row 3.
    digestion_data : pd.DataFrame, optional
        DataFrame containing pre-loaded 'mass_mg' and 'volume_ml' to bypass manual entry.

    Returns:
    --------
    tuple (pd.DataFrame, pd.DataFrame)
        A tuple containing (df_final_pct, df_final_ppm).
    """
    df = df_clean_samples.copy()

    # Separate metadata (first 2 rows) from the sample data block
    headers = df.iloc[:2].copy()
    data = df.iloc[2:].copy()

    sample_col = data.columns[0]
    sample_list = data[sample_col].tolist()

    masses_mg = []
    volumes_ml = []

    # Bypasses CLI input if structured digestion data is provided
    if digestion_data is not None and isinstance(digestion_data, pd.DataFrame):
        s_masses = digestion_data['mass_mg']
        s_volumes = digestion_data['volume_ml']
    else:
        print("\n=======================================================")
        print(" DIGESTION DATA ENTRY (MASS AND DILUTION VOLUME)")
        print("=======================================================")

        for sample in sample_list:
            print(f"\n--- Sample: {sample} ---")

            # Input Sample Mass (mg)
            while True:
                try:
                    m_str = input(f" > Sample mass for '{sample}' (mg): ").replace(',', '.')
                    m_val = float(m_str)
                    if m_val <= 0:
                        print("    ⚠️ Mass must be greater than zero.")
                        continue
                    masses_mg.append(m_val)
                    break
                except ValueError:
                    print("    ⚠️ Please enter a valid number.")

            # Input Dilution Volume (mL or g)
            while True:
                try:
                    v_str = input(f" > Dilution volume/weight for '{sample}' (mL or g): ").replace(',', '.')
                    v_val = float(v_str)
                    if v_val <= 0:
                        print("    ⚠️ Dilution volume or weight must be greater than zero.")
                        continue
                    volumes_ml.append(v_val)
                    break
                except ValueError:
                    print("    ⚠️ Please enter a valid number.")

        s_masses = pd.Series(masses_mg, index=data.index)
        s_volumes = pd.Series(volumes_ml, index=data.index)

    # Analytical element columns (readings in ppb = µg/L)
    element_cols = data.columns[1:]

    df_ppm = pd.DataFrame(index=data.index)
    df_pct = pd.DataFrame(index=data.index)

    # Conversion Factor: 1% = 10,000 ppm
    FACTOR_PPM_TO_PCT = 10000.0

    for col in element_cols:
        # Convert values to numeric, replace commas, and handle missing entries
        icp_reading = pd.to_numeric(
            data[col].astype(str).str.replace(',', '.'), 
            errors='coerce'
        ).fillna(0.0)

        # 1. Calculate Concentration in ppm (mg/kg)
        conc_ppm = (icp_reading * s_volumes) / s_masses

        # 2. Convert Concentration to Percentage (%)
        conc_pct = conc_ppm / FACTOR_PPM_TO_PCT

        df_ppm[col] = conc_ppm
        df_pct[col] = conc_pct

    # Reinsert the sample names column
    df_ppm.insert(0, sample_col, data[sample_col])
    df_pct.insert(0, sample_col, data[sample_col])

    # Reconstruct final DataFrames incorporating initial metadata headers
    df_final_ppm = pd.concat([headers.copy(), df_ppm], ignore_index=True)
    df_final_pct = pd.concat([headers.copy(), df_pct], ignore_index=True)

    print("\n✓ Unit conversions and concentration calculations completed successfully.")
    return df_final_pct, df_final_ppm
