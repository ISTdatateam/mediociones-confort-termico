import pandas as pd

def get_data(csv_url: str) -> pd.DataFrame:
    """
    Lee los datos desde un CSV remoto (o local) y retorna un DataFrame de pandas.
    """
    df = pd.read_csv(csv_url)
    return df
