# IVN_ISM_Simulation
# IVN-ISM and Fuzzy ISM Simulation

This repository contains a VBA implementation of a simulation designed to validate the **Interval-Valued Neutrosophic Interpretive Structural Modeling (IVN-ISM)** methodology by comparing it with the classical **Fuzzy ISM** approach. The validation uses the **Dice-Sørensen similarity index** to assess the similarity between the results of the two methodologies.

---

## Features

- Generates random expert opinions and decision matrices.
- Constructs reachability matrices for both IVN-ISM and Fuzzy ISM.
- Converts decision matrices into binary matrices using thresholding.
- Computes the **Dice-Sørensen similarity index** to measure structural similarity.
- Supports multiple replications to ensure robust statistical analysis.

---

## Simulation Details

- **Input:** Number of replications (specified by the user at runtime).
- **Output:**
  - Average Dice-Sørensen similarity index.
  - Standard deviation of similarity values.
  - Results are written directly into an Excel worksheet.

---

## Requirements

- Microsoft Excel with VBA enabled.
- Basic understanding of VBA for customization (optional).

---

## Installation

1. Open Microsoft Excel.
2. Press `Alt + F11` to open the VBA editor.
3. Insert a new module by navigating to `Insert > Module`.
4. Copy the code from the file `IVN_ISM_Simulation.vba` and paste it into the module.
5. Save the workbook as a macro-enabled file (`.xlsm`).

---

## Usage

1. Run the macro named `IVN_ISM_Simulation` by pressing `Alt + F8` in Excel.
2. Enter the number of replications when prompted.
3. The simulation will calculate:
   - The average Dice-Sørensen similarity index.
   - The standard deviation of the similarity index.
4. Results will be displayed in the active worksheet:
   - Cell `C2`: Average Dice-Sørensen similarity index.
   - Cell `C3`: Standard deviation of similarity values.

---

## Example Output

| Metric                              | Value  |
|-------------------------------------|--------|
| Average Dice-Sørensen Similarity    | 0.8161 |
| Standard Deviation                  | 0.0417 |

---

## Explanation of the Dice-Sørensen Similarity Index

The **Dice-Sørensen similarity index** is a statistical measure used to evaluate the overlap between two binary datasets. It is defined as:

\[
\text{Dice-Sørensen Index} = \frac{2 \times |\text{Intersection of Sets}|}{|\text{Set A}| + |\text{Set B}|}
\]

This index is commonly used for measuring structural similarity in decision matrices.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Author

- **[Your Name]**  
Feel free to reach out for questions or suggestions regarding the project.
