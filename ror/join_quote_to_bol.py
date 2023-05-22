import pandas as pd
import polars as pl


# Filenames
bol_filename = "ROR_BillofLading.xlsx"
workorder_filename = "ROR_WorkOrders.xlsx"
quote_filename = "ROR_Quotes.xlsx"

# Load  data
bol_df = pl.from_pandas(pd.read_excel(bol_filename))
workorder_df = pl.from_pandas(pd.read_excel(workorder_filename))
quote_df = pl.from_pandas(pd.read_excel(quote_filename))

# Create lookup table
bol_workorder_junction = (
    bol_df
        .select(["BOLNumber", "SelectWorkOrder"])
        .rename({
            "BOLNumber": "BOLNumber",
            "SelectWorkOrder": "WorkOrderNumber",
        })
)
workorder_quote_junction = (
    workorder_df
        .select(["WorkOrderNumber", "SelectQuote"])
        .rename({
            "WorkOrderNumber": "WorkOrderNumber",
            "SelectQuote": "QuoteNumber",
        })
)
lookup_number_df = bol_workorder_junction.join(workorder_quote_junction, on="WorkOrderNumber")

# Calculate stats
prices_df = lookup_number_df.join(
    quote_df.select([
        pl.col("QuoteNumber"),
        pl.col("SP_PricePer_Load"),
        pl.col("SP_PricePer_Lineal"),
        pl.col("SP_PricePer_Square"),
        pl.col("SP_CompactionCost"),
        pl.col("SP_MeasuringCost"),
    ]),
    on="QuoteNumber",
)
prices_df = bol_df.join(prices_df, on="BOLNumber")
    
prices_df.write_excel("ROR_BillofLading_AddedColumns.xlsx")
