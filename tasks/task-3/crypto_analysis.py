import os
import shutil
import win32com.client

def build_dashboard():
    """
    Reads the dataset from crypto_dataset.xlsx and produces
    a separate output file crypto_analysis_output.xlsx with
    the Price Change Dashboard (PivotTables, Slicer, Chart).
    The original dataset file is NOT modified.
    """

    dataset_path = r"D:\task-3\crypto_dataset.xlsx"
    output_path  = r"D:\task-3\crypto_analysis_output.xlsx"

    if not os.path.exists(dataset_path):
        raise FileNotFoundError(f"Dataset not found: {dataset_path}")

    # Copy the dataset to create an independent output file
    shutil.copy2(dataset_path, output_path)
    print(f"Copied dataset -> {output_path}")

    # Launch Excel
    print("Launching Excel...")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        # Open ONLY the output copy
        wb = excel.Workbooks.Open(output_path)

        # Remove every sheet except 'Crypto Data'
        for sh in list(wb.Sheets):
            if sh.Name != "Crypto Data":
                name = sh.Name
                sh.Delete()
                print(f"  Deleted sheet: {name}")

        # Create the dashboard sheet
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws.Name = "Price Change Analysis"
        print("Created sheet: Price Change Analysis")

        ws.Activate()
        excel.ActiveWindow.DisplayGridLines = True

        # Title
        ws.Cells(2, 1).Value = "Price Change Dashboard"
        ws.Cells(2, 1).Font.Size = 16
        ws.Cells(2, 1).Font.Bold = True
        ws.Cells(2, 1).Font.Color = 0x4F81BD

        # ── PivotCache ──────────────────────────────────────
        src = "'Crypto Data'!$A$1:$R$201"
        pc  = wb.PivotCaches().Create(SourceType=1, SourceData=src)

        # ── PT1: Symbol + Price Increase table ──────────────
        print("Creating PivotTable 1 (Symbol & Price Change)...")
        pt1 = pc.CreatePivotTable(
            TableDestination="'Price Change Analysis'!R5C1",
            TableName="PT_SymbolChange")
        pt1.ColumnGrand = False
        pt1.RowGrand    = False

        sf1 = pt1.PivotFields("Symbol")
        sf1.Orientation = 1   # xlRowField
        sf1.Position    = 1

        vf1 = pt1.PivotFields("Price Increased (1h)")
        vf1.Orientation = 4   # xlDataField
        vf1.Function    = -4157  # xlSum
        vf1.Name        = "Price Increase"

        sf1.PivotFilters.Add2(
            Type=5, DataField=pt1.PivotFields("Price Increase"), Value1=10)
        sf1.AutoSort(2, "Price Increase")   # descending

        # ── PT2: Chart data (Current Price + Previous Price) ─
        print("Creating PivotTable 2 (Chart data)...")
        pt2 = pc.CreatePivotTable(
            TableDestination="'Price Change Analysis'!R5C5",
            TableName="PT_ChartData")
        pt2.ColumnGrand = False
        pt2.RowGrand    = False

        sf2 = pt2.PivotFields("Symbol")
        sf2.Orientation = 1
        sf2.Position    = 1

        for col, alias in [("Price (USD)",            "Current Price"),
                           ("1h Before Price (USD)",  "Previous Price"),
                           ("Price Increased (1h)",   "Increase")]:
            f = pt2.PivotFields(col)
            f.Orientation = 4
            f.Function    = -4157
            f.Name        = alias

        sf2.PivotFilters.Add2(
            Type=5, DataField=pt2.PivotFields("Increase"), Value1=10)
        sf2.AutoSort(2, "Increase")

        # ── Slicer: Price Category ───────────────────────────
        print("Creating Slicer...")
        try:
            wb.SlicerCaches("Slicer_PriceCat").Delete()
        except Exception:
            pass

        sc = wb.SlicerCaches.Add2(
            Source=pt1, SourceField="Price Category",
            Name="Slicer_PriceCat")
        sc.Slicers.Add(
            SlicerDestination=ws,
            Name="SlicerPriceCat",
            Caption="Price Range Category",
            Left=700, Top=50, Width=160, Height=100)
        sc.PivotTables.AddPivotTable(pt2)   # link to PT2 too

        # ── Chart ────────────────────────────────────────────
        print("Creating Chart...")
        co = ws.ChartObjects().Add(Left=320, Top=170, Width=520, Height=300)
        ch = co.Chart
        ch.SetSourceData(Source=pt2.TableRange1)
        ch.ChartType  = 51        # xlColumnClustered
        ch.HasTitle   = True
        ch.ChartTitle.Text = "Top 10 Price Change (Present vs 1h Before)"
        ch.HasLegend  = True
        ch.Legend.Position = -4107  # bottom

        try:
            ch.SeriesCollection("Increase").Delete()
            print("  Hidden 'Increase' series from chart.")
        except Exception:
            pass

        ws.Columns("A:C").AutoFit()
        ws.Columns("E:H").AutoFit()

        # ── Save ─────────────────────────────────────────────
        wb.Save()
        wb.Close()
        print(f"\nOutput saved: {output_path}")
        print(f"Dataset kept: {dataset_path}")

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

    finally:
        excel.Quit()
        print("Done.")


if __name__ == "__main__":
    build_dashboard()
