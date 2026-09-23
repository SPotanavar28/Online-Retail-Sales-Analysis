import nbformat as nbf
from nbclient import NotebookClient
import time
import os

input_nb_path = "SnehaVinodPotanavar_OnlineRetailSalesAnalysis.ipynb"

print(f"Reading {input_nb_path}...")
with open(input_nb_path, "r", encoding="utf-8") as f:
    nb = nbf.read(f, as_version=4)

print(f"Total cells to execute: {len(nb.cells)}")
print("Executing notebook with NotebookClient...")
start_time = time.time()

client = NotebookClient(nb, timeout=600, kernel_name="python3")
client.execute()

output_nb_path = "SnehaVinodPotanavar_OnlineRetailSalesAnalysis.ipynb"
with open(output_nb_path, "w", encoding="utf-8") as f:
    nbf.write(nb, f)

elapsed = time.time() - start_time
print(f"Execution completed in {elapsed:.2f} seconds!")
print(f"Executed notebook saved to {output_nb_path} with all outputs and inline charts intact.")
