import win32com.client as win32
import os
import shutil


def stop_excel():
    try:
        os.system("taskkill /F /IM EXCEL.EXE")
    except Exception as _:
        pass


def converter_xls_para_xlsx(
    input_file: str, output_file: str, deletar_arquivo: bool = False
):
    """Essa funcao abre o aplicativo excel e cria e transforma o xls em xlsx"""
    file_name = "convertendo_excel.xls"

    temp = os.path.join(os.path.dirname(__file__), "temp")
    excel_temp = os.path.join(temp, file_name)  # Caminho do arquivo xls temporario
    excel_temp_xlsx = os.path.join(temp, file_name + "x")

    if os.path.exists(excel_temp_xlsx):
        os.remove(excel_temp_xlsx)

    shutil.copy(input_file, excel_temp)
    excel_conv = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        wb_file = excel_conv.Workbooks.Open(excel_temp)
        try:
            wb_file.SaveAs(
                excel_temp_xlsx, FileFormat=51
            )  # FileFormat: 51 == .xlsx | 56 == .xls
        finally:
            wb_file.Close()
    finally:
        excel_conv.Application.Quit()

    if os.path.isfile(output_file):
        os.remove(output_file)
    if os.path.isfile(excel_temp_xlsx):
        os.rename(
            excel_temp_xlsx, output_file
        )  # Renomeando arquivo com o nome do output_file
    if os.path.isfile(excel_temp):
        os.remove(excel_temp)
    if deletar_arquivo and os.path.isfile(input_file + "x"):
        os.remove(input_file)
