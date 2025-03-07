# Sleep-Wake Transitions

This repository contains code for analyzing **sleep-wake transitions in rodents** using **piezoelectric sleep cages and SleepStats Data Explorer Version 4**.

## Files
- `CalculateSleepWakeTransitions.bas` - **VBA Macro** for batch processing of sleep-wake transitions from SleepStats outputs.
- `Mannino_Transitions_Paper.pdf` - **Paper detailing methodology and validation**.
- `LICENSE` - **MIT License** for open use.

## SleepStats Integration
- The dataset should be **exported from SleepStats Version 4** as a CSV file.
- Follow the **data processing steps outlined in the Mannino et al. paper** for proper formatting.

## Usage
- Import CSV files from **SleepStats Data Explorer**.
- Use `CalculateSleepWakeTransitions.bas` in **Excel VBA** to analyze transitions.
- Review **transition outputs per hour**, consistent with the methodology in **Mannino et al.**.

**Note:** This repository is private until publication.
