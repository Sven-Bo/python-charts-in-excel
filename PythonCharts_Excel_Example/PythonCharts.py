from pathlib import Path

import matplotlib.pyplot as plt  # pip install matplotlib
import mplfinance as mpf  # pip install mplfinance
import pandas as pd  # pip install pandas
import plotly.express as px  # pip install plotly-express
import seaborn as sns  # pip install seaborn
import xlwings as xw  # pip install xlwings
import yfinance as yf  # pip install yfinance



def display_msgbox(wb, text):
    msg_box = wb.macro("Module1.msg_box")
    msg_box(text)
    return None


def insert_picture_to_excel(sht, cell, fig, pic_name):
    sht.pictures.add(
        fig,
        name=pic_name,
        update=True,
        left=sht.range(cell).left,
        top=sht.range(cell).top,
        height=200,
        width=300,
    )
    return None


def matplotlib_bar_chart(df):
    fig = plt.figure()
    x = df["day"]
    y = df["total_bill"]
    plt.bar(x, y)
    plt.grid(False)
    plt.ylabel("in USD")
    plt.title("Total Bill Amount By Day")
    return fig


def pandas_bar_chart(df):
    df_grouped = df.groupby(by="day", as_index=False).sum()
    ax = df_grouped.plot(kind="bar", x="day", y="tip", color="#50C878", grid=False)
    fig = ax.get_figure()
    return fig


def seaborn_bar_chart(df):
    fig = plt.figure()
    sns.set_style({"axes.grid": False})
    sns.barplot(data=df, x="day", y="total_bill", hue="sex", ci=None)
    return fig


def seaborn_scatter_plot(df):
    fig = plt.figure()
    sns.scatterplot(data=df, x="total_bill", y="tip", hue="day", style="time")
    return fig


def plotly_histogram(df):
    fig = px.histogram(df, x="day", y="total_bill", color="sex")
    return fig


def plotly_histogram_advanced(df):
    fig = px.histogram(
        df,
        x="day",
        y="total_bill",
        color="sex",
        title="Receipts by Payer Gender and Day of Week vs Target",
        labels={"sex": "Payer Gender", "day": "Day of Week", "total_bill": "Receipts"},
        category_orders={
            "day": ["Thur", "Fri", "Sat", "Sun"],
            "sex": ["Male", "Female"],
        },
        color_discrete_map={"Male": "RebeccaPurple", "Female": "MediumPurple"},
        template="simple_white",
    )

    fig.update_yaxes(tickprefix="$", showgrid=True)  # the y-axis is in dollars

    fig.update_layout(  # customize font and legend orientation & position
        font_family="Rockwell",
        legend=dict(
            title=None, orientation="h", y=1, yanchor="bottom", x=0.5, xanchor="center"
        ),
    )

    fig.add_shape(  # add a horizontal "target" line
        type="line",
        line_color="salmon",
        line_width=3,
        opacity=1,
        line_dash="dot",
        x0=0,
        x1=1,
        xref="paper",
        y0=950,
        y1=950,
        yref="y",
    )

    fig.add_annotation(  # add a text callout with arrow
        text="below target!", x="Fri", y=400, arrowhead=1, showarrow=True
    )
    return fig


def create_candle_chart(ticker, start, end, OUTPUT_DIR):
    # Get stock data
    ticker = ticker
    start = start
    end = end
    data = yf.download(ticker, start=start, end=end)

    # Create candle stick chart & save to output dir
    output_img = OUTPUT_DIR / "mplfiance_candle.png"
    mpf.plot(
        data,
        type="candle",
        volume=True,
        style="yahoo",
        axtitle=f"{ticker}",
        savefig=output_img,
    )
    return output_img


def create_line_chart(ticker, start, end, OUTPUT_DIR):
    # Get stock data
    ticker = ticker
    start = start
    end = end
    data = yf.download(ticker, start=start, end=end)

    # Create candle stick chart & save to output dir
    output_img = OUTPUT_DIR / "mplfiance_line.png"
    mpf.plot(
        data, type="line", style="yahoo", axtitle=f"{ticker}", savefig=output_img
    )
    return output_img


def plot_tips_data():
    wb = xw.Book.caller()
    sht = wb.sheets["Python Charts"]
    df = sht.range("A1").options(pd.DataFrame, index=False, expand='table').value

    # Generate & insert charts into Excel
    fig = matplotlib_bar_chart(df)
    insert_picture_to_excel(sht=sht, cell="I2", fig=fig, pic_name="Matplotlib")
    fig = pandas_bar_chart(df)
    insert_picture_to_excel(sht=sht, cell="I15", fig=fig, pic_name="Pandas")
    fig = seaborn_bar_chart(df)
    insert_picture_to_excel(sht=sht, cell="N2", fig=fig, pic_name="SeabornBarChart")
    fig = seaborn_scatter_plot(df)
    insert_picture_to_excel(sht=sht, cell="N15", fig=fig, pic_name="SeabornScatterPlot")
    fig = plotly_histogram(df)
    insert_picture_to_excel(sht=sht, cell="I28", fig=fig, pic_name="PlotlyHistogram")
    fig = plotly_histogram_advanced(df)
    insert_picture_to_excel(
        sht=sht, cell="N28", fig=fig, pic_name="PlotlyHistogramAdvanced"
    )

    display_msgbox(wb, "Task completed!")


def plot_stock_data():
    wb = xw.Book.caller()
    sht = wb.sheets["Stock Dashboard"]

    # Create output directory
    OUTPUT_DIR = Path(__file__).parent / "Output"
    OUTPUT_DIR.mkdir(exist_ok=True)

    # Get stock input
    ticker = sht.range("TICKER").value
    start = sht.range("START_DATE").value
    end = sht.range("END_DATE").value

    # Generate & insert charts into Excel
    fig = create_candle_chart(ticker, start, end, OUTPUT_DIR)
    insert_picture_to_excel(sht=sht, cell="C5", fig=fig, pic_name="CandleChart")
    fig = create_line_chart(ticker, start, end, OUTPUT_DIR)
    insert_picture_to_excel(sht=sht, cell="C18", fig=fig, pic_name="LineChart")

    display_msgbox(wb, "Task completed!")


if __name__ == "__main__":
    xw.Book("PythonCharts.xlsm").set_mock_caller()
