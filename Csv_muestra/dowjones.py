
#%%
import pandas as pd
import seaborn as snb
from pathlib import Path
import matplotlib.pyplot as plt

ROOT = Path(__file__).resolve().parent


df = pd.read_csv(ROOT / "dowjones.csv")

print(df.columns)

snb.lineplot(data=df, x="Date", y="Price")
plt.title("Evolución histórica del Dow Jones")
plt.xlabel("Fecha")
plt.ylabel("Precio")
plt.tight_layout()
plt.show()
# %%
