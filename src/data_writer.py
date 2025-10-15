import pandas as pd
import os
import json
from .logger import logger

class DataWriter:
    def __init__(self, config_path='config/config.json'):
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        self.fields = self.config['fields']

    def write_data(self, data_list, output_file):
        """Write extracted data to file (Excel or CSV)."""
        if not data_list:
            logger.warning("No data to write")
            return 0

        df = pd.DataFrame(data_list)
        df = df.reindex(columns=self.fields).fillna("No encontrado")

        ext = os.path.splitext(output_file)[1].lower()
        try:
            if ext == '.xlsx':
                self._write_excel(df, output_file)
            elif ext == '.csv':
                self._write_csv(df, output_file)
            else:
                raise ValueError(f"Unsupported output format: {ext}")
            logger.info(f"Data written to {output_file} with {len(df)} records")
            return len(df)
        except Exception as e:
            logger.error(f"Error writing to {output_file}: {e}")
            raise

    def _write_excel(self, df, output_file):
        """Write to Excel, appending if exists."""
        if os.path.exists(output_file):
            try:
                df_existing = pd.read_excel(output_file)
                df_combined = pd.concat([df_existing, df], ignore_index=True)
                df_combined.to_excel(output_file, index=False)
                logger.info(f"Appended to existing Excel file {output_file}")
            except Exception as e:
                logger.warning(f"Could not read existing Excel, overwriting: {e}")
                df.to_excel(output_file, index=False)
        else:
            df.to_excel(output_file, index=False)

    def _write_csv(self, df, output_file):
        """Write to CSV, appending if exists."""
        mode = 'a' if os.path.exists(output_file) else 'w'
        header = not os.path.exists(output_file)
        df.to_csv(output_file, index=False, mode=mode, header=header, encoding='utf-8')
        if mode == 'a':
            logger.info(f"Appended to existing CSV file {output_file}")