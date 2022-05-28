import pandas as pd
import matplotlib
from pylab import title, figure, xlabel, ylabel, xticks, bar, legend, axis, savefig
from fpdf import FPDF
import os


def GeneratePDF():

    dataFrame = pd.read_excel(os.path.dirname(__file__) + "/20220201-143409_UserName_24_capture_xlsx.xlsx")
    """
    dataFrame['ID'] = id
    dataFrame['HR'] = hr
    dataFrame['RRInterval'] = rrInterval
    dataFrame['Temperature'] = temperature
    dataFrame['AccelerometerX'] = accelerometerX
    dataFrame['AccelerometerY'] = accelerometerY
    dataFrame['AccelerometerZ'] = accelerometerZ
    dataFrame['GyroX'] = gyroX
    dataFrame['GyroY'] = gyroY
    dataFrame['GyroZ'] = gyroZ
    """
    #file_name = export_dir + r'/'+strMetadata+'capture_xlsx.xlsx'

    print(dataFrame.describe())


    df = pd.DataFrame()
    df['Question'] = ["Q1", "Q2", "Q3", "Q4"]
    df['Charles'] = [3, 4, 5, 3]
    df['Mike'] = [3, 3, 4, 4]

    title("Monitoramento parâmetros fisiológicos")

    xlabel('Question Number')
    ylabel('Score')
    c = [2.0, 4.0, 6.0, 8.0]
    m = [x - 0.5 for x in c]
    xticks(c, df['Question'])
    bar(m, df['Mike'], width=0.5, color="#91eb87", label="Mike")
    bar(c, df['Charles'], width=0.5, color="#eb879c", label="Charles")
    legend()
    axis([0, 10, 0, 8])
    savefig('barchart.png')

    pdf = FPDF()
    pdf.add_page()
    pdf.set_xy(0, 0)
    # Page Lineborders
    pdf.set_line_width(0.0)
    pdf.line(5.0,5.0,205.0,5.0) # top one
    pdf.line(5.0,292.0,205.0,292.0) # bottom one
    pdf.line(5.0,5.0,5.0,292.0) # left one
    pdf.line(205.0,5.0,205.0,292.0) # right one


    pdf.set_font('arial', 'B', 12)
    pdf.cell(20)
    #pdf.cell(75, 10, "", 1, 2, 'C')
    pdf.cell(75, 10, "Monitoramento parâmetros fisiológicos", 1, 2, 'C')

    # Blank line
    pdf.cell(ln=1, h=5.0, align='L', w=0, txt="", border=0)



    pdf.set_font('arial', 'B', 10)
    pdf.cell(20, 10, "Informação geral", 0, 2, 'L')
    pdf.set_font('arial', 'B', 8)
    pdf.cell(20, 10, "Data e horário de avaliação:", 0, 2, 'L')
    pdf.cell(20, 10, "Mergulhador avaliado:", 0, 2, 'L')
    pdf.cell(20, 10, "Idade:", 0, 2, 'L')

    #pdf.cell(20, 10, 'Title', 1, 1, 'C')

    pdf.cell(90, 10, " ", 0, 2, 'C')
    pdf.cell(-40)

    pdf.cell(50, 10, 'Question', 1, 0, 'C')
    pdf.cell(40, 10, 'Charles', 1, 0, 'C')
    pdf.cell(40, 10, 'Mike', 1, 2, 'C')
    pdf.cell(-90)
    pdf.set_font('arial', '', 12)
    for i in range(0, len(df)):
        pdf.cell(50, 10, '%s' % (df['Question'].iloc[i]), 1, 0, 'C')        
        pdf.cell(40, 10, '%s' % (str(df.Mike.iloc[i])), 1, 0, 'C')
        pdf.cell(40, 10, '%s' % (str(df.Charles.iloc[i])), 1, 2, 'C')
        pdf.cell(-90)

    pdf.cell(90, 10, " ", 0, 2, 'C')
    pdf.cell(-30)

    pdf.image('barchart.png', x=None, y=None, w=0, h=0, type='', link='')
    pdf.output('test.pdf', 'F')

if __name__ == "__main__":
    GeneratePDF()