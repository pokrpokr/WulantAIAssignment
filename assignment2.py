import camelot
import click
import traceback
import os
import fitz
from openpyxl import Workbook
    
def process_pdf(src_pdf, expt_name, page_start, page_end):
    doc = fitz.Document(src_pdf)
    if page_start == 0 and page_start == page_end:
        pages = "all" 
    elif page_start > 0 and page_start < len(doc) and page_start == page_end:
        pages = f"{page_start}"
    elif page_start > 0 and page_start < len(doc) and page_end > page_start and page_end <= len(doc):
        pages = f"{page_start}-{page_end}"
    else:
        raise Exception("Can not find pages in PDF")
    doc.close()
   
    tables = camelot.read_pdf(src_pdf, flavor='stream', pages=pages, edge_tol=50, scale=60) # edge_tol=500, flag_size=True, 
    print("Total tables extracted:", tables.n)

    if tables.n > 0:
        wb = Workbook()
        wb.remove(wb.active)
        for table in tables:
            # camelot.plot(table, kind='contour').show()
            df = table.df
            ws = wb.create_sheet(title=f"Page{table.parsing_report['page']} Tables")

            ws.append(df.columns.to_list())

            for row in df.itertuples(index=False):
                ws.append(row)
        wb.save(expt_name)
    else:
        print("No tables in selected pages")
    
@click.command()
@click.option("--src_pdf", "-sp", type=str, prompt="Please enter PDF file name")
@click.option("--expt_name", "-en", type=str, prompt="Please enter output excel file name")
@click.option("--ver", "-ver", type=int, default=1, prompt="ver: 1 using Camelot-py; ver: 2 using Cascadetabnet")
@click.option("--pg_start", "-pgs", type=int, default=0, prompt="PDF start page")
@click.option("--pg_end", "-pge", type=int, default=0, prompt="PDF end page")
def main(src_pdf, expt_name, ver, pg_start, pg_end):
    if ver == 1:
        process_pdf(os.path.join(os.getcwd()+"/test_pdfs/", src_pdf), os.path.join(os.getcwd()+"/export_excels/", expt_name), pg_start, pg_end)
        print(f"----- File location: {os.path.join(os.getcwd()+'/export_excels/', expt_name)} -----")
    else:
        print("No other versions")
if __name__ == "__main__":
    try:
        print("----- Processiing PDF -----")
        main()
        print("----- Export Finished -----")
    except Exception as e:
        print(traceback.format_exc())
        print(f"Error occur! Msg: {str(e)}")
        
