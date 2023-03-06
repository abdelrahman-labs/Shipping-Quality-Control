import pandas as pd
import streamlit as st
import plotly.express as px
from st_aggrid import GridOptionsBuilder, AgGrid
import plotly.graph_objects as go
import re
import numpy as np
import datetime
import pydeck as pdk
import os

st.set_page_config(page_title="Quality Control", page_icon=":bar_chart:", layout="wide")

@st.cache(allow_output_mutation=True)
def dataframe():
    data = pd.read_excel("Agency (G).xlsx", sheet_name="Raw_Data")
    data[[col for col in data.columns if 'time' in col or 'Time' in col]] = data[[col for col in data.columns if 'time' in col or 'Time' in col]].apply(pd.to_datetime)
    data[["Multiple Deliveries", "Count of Abnormals", 'Count of "No Answer" Abnormal', 'no update', "Auto Sign Shipments (Critical!)"]] = data[
        ["Multiple Deliveries", "Count of Abnormals", 'Count of "No Answer" Abnormal', 'no update', "Auto Sign Shipments (Critical!)"]].astype(float)
    data.loc[(data["Days of stay"] == "More than 7 Days") | (data["Days of stay"] == "More than 10 Days (Lost?)"), "exceeded"] = "Exceeded"
    data.loc[(data["Days of stay"] == "More than 6 Days (Critical!)"), "exceeded"] = "About to Exceed"
    data.loc[(data["exceeded"].isna()), "exceeded"] = "Not Exceed"
    data.loc[(data["Days of stay"] == "More than 7 Days") | (data["Days of stay"] == "More than 10 Days (Lost?)"), "Flag1"] = "Exceeded life cycle"
    data.loc[(data["Delay Reason"] == "No Return Application"), "Flag2"] = "No Return Application"
    data.loc[(data["Print status"] == "Unprinted") & (data["Application Status"] == "Reviewed"), "Flag3"] = "Return applications to be printed"
    data.loc[(data["Multiple Deliveries"].isna()), "Flag4"] = "Didn't get out for delivery"
    data.loc[(data['Count of "No Answer" Abnormal'] > 1), "Flag5"] = 'More than 1 "No Answer" abnormal'
    data.loc[(data['no update'] > 1), "Flag6"] = "No Update since More Than 24 Hours"
    data['Action'] = data[['Flag1', 'Flag2', 'Flag3', 'Flag4', 'Flag5', 'Flag6']].fillna('').agg(','.join, axis=1)
    data['Action'] = data['Action'].apply(lambda x: [o for o in x.split(',') if len(o)])
    data.drop(['Flag1', 'Flag2', 'Flag3', 'Flag4', 'Flag5', 'Flag6'], axis=1, inplace=True)
    data["confirmdate"] = [d.date() for d in data["Auditing Time"]]
    data["todaydate"] = datetime.datetime(2023, 3, 5).date()
    data.loc[(data["Print status"] == "Unprinted") & (data["Application Status"] == "Reviewed") & (data["confirmdate"].notna()) & ((data["todaydate"] - data["confirmdate"]).dt.days == 1), "unprintedold"] = True
    data.loc[(data["Print status"] == "Unprinted") & (data["Application Status"] == "Reviewed") & (data["confirmdate"].notna()) & ((data["todaydate"] - data["confirmdate"]).dt.days > 1), "unprintedolder"] = True
    data.loc[(data["Print status"] == "Unprinted") & (data["Application Status"] == "Reviewed") & (data["confirmdate"].notna()) & ((data["todaydate"] - data["confirmdate"]).dt.days == 0), "unprintednew"] = True
    data.loc[data["All Abnormal Sequence (Old to New)"].isna(), "All Abnormal Sequence (Old to New)"] = 0
    data["printdate"] = [d.date() for d in data["Print Time"]]
    data["todaydate"] = datetime.datetime.today().date()
    data.loc[data['printdate'].notna(), "todayminusprint"] = (data.loc[(data['printdate'].notna()), "todaydate"] - data.loc[(data['printdate'].notna()), "printdate"]).dt.days
    data.loc[data['todayminusprint'] == 0, "pnmnew"] = True
    data.loc[data['todayminusprint'] == 1, "pnmold"] = True
    data.loc[data['todayminusprint'] > 1, "pnmolder"] = True

    lastupdated = pd.read_excel("Agency (G).xlsx", sheet_name="Last_Updated", header=None).iloc[0,0]
    unpickup = pd.read_excel("Agency (G).xlsx", sheet_name="Unpickup")
    for i in unpickup.columns:
        try:
            unpickup[i] = unpickup[i].astype(int)
        except (TypeError, ValueError):
            pass
    unpickup['Date'] = unpickup['Date'].dt.date
    on_time = pd.read_excel("Agency (G).xlsx", sheet_name="on-time")
    on_time[[col for col in on_time.columns if 'Amount' in col]] = on_time[[col for col in on_time.columns if 'Amount' in col]].astype(float)
    on_time = on_time.sort_values(by="Date", ascending=False)
    return data, lastupdated, on_time, unpickup


df, lastupdated, ontimedf, unpickupdf = dataframe()
tabb1, tabb2 = st.tabs(["Unsigned Shipments", "Couriers Monitoring"])

st.write(unpickupdf)
with tabb1:
    # Sidebar
    st.sidebar.title("Filter Options")

    # Agency filter
    agency_options = df.sort_values(by="Agency")["Agency"].unique()
    # default_agency = [agency_options[0]] if len(agency_options) > 0 else []
    agency = st.sidebar.multiselect("Select Agency:", options=agency_options, default=agency_options)

    # Branch/DC filter
    branch_options = df.sort_values(by="Latest Scan Branch").query("Agency == @agency")["Latest Scan Branch"].unique()
    # default_branch = [branch_options[0]] if len(branch_options) > 0 else []
    branch = st.sidebar.multiselect("Select Branch/DC:", options=branch_options, default=branch_options)

    # agency = st.sidebar.multiselect("Select Agency:", options=df["Agency"].unique())
    # branch = st.sidebar.selectbox("Select Branch/DC:", options=df.query("Agency == @agency")["Latest Scan Branch"].unique())
    df_selection = df.query('Agency == @agency & `Latest Scan Branch` == @branch')
    ontimedf_selection = ontimedf.query('`Delivery Branch Name` == @branch')
    ontimedf_selection = ontimedf_selection.groupby("Date")[['Receivable Amount', 'On-time signing Amount']].sum().reset_index().sort_values(by="Date", ascending=False)
    unpickupdf_selection = unpickupdf.query('Branch == @branch')
    unpickupdf_selection = unpickupdf_selection.groupby("Date")['Total Unpick-ip'].sum().reset_index().sort_values(by="Date", ascending=True)
    unpickupdf_selection = unpickupdf_selection.loc[unpickupdf_selection["Total Unpick-ip"] > 0]
    unpickups = unpickupdf_selection["Total Unpick-ip"].sum()

    latest_ontime = ontimedf.groupby(["Date", "Agency Area Name", "Delivery Branch Name"])[['Receivable Amount', 'On-time signing Amount']].sum().reset_index().sort_values(by="Date", ascending=False).loc[
        ontimedf.groupby(["Date", "Delivery Branch Name"])[['Receivable Amount', 'On-time signing Amount']].sum().reset_index().sort_values(by="Date", ascending=False)["Date"] ==
        ontimedf.groupby(["Date", "Delivery Branch Name"])[['Receivable Amount', 'On-time signing Amount']].sum().reset_index().sort_values(by="Date", ascending=False).iloc[0, 0]]
    latest_ontime["On-Time Sign Rate"] = ((latest_ontime["On-time signing Amount"] / latest_ontime['Receivable Amount']) * 100)
    latest_ontime = latest_ontime.query('`Delivery Branch Name` == @branch')

    # mainpage
    title, gauge = st.columns(2)
    try:
        on_time = ontimedf_selection.iloc[0, 2] / ontimedf_selection.iloc[0, 1]
        yesterday_ontime = ontimedf_selection.iloc[1, 2] / ontimedf_selection.iloc[1, 1]
        ontimedate = str(ontimedf_selection.iloc[0, 0])[:10]
    except IndexError:
        pass

    with title:
        st.title(":bar_chart: Quality Control")
        st.write(f"Last updated: {lastupdated}")
        if st.button("Hide All"):
            st.empty()
        st.markdown("""<style>.big-font {font-size:20px !important;}</style>""", unsafe_allow_html=True)
        st.markdown(
            '<p class="big-font">This On-Time Chart is updated daily. The Arrow under the on-time rate is a comparison with the previous day.</p>',
            unsafe_allow_html=True)


        st.markdown("##")

    with gauge:
        try:
            gauge = go.Figure(go.Indicator(
                mode="gauge+number+delta",
                number={'valueformat': ".2%", 'font': {'size': 60}},
                value=on_time,
                domain={'x': [0, 1], 'y': [0, 1]},
                title={'text': f"On-Time Sign Rate\n{ontimedate}", 'font': {'size': 25}},
                delta={'reference': yesterday_ontime, 'valueformat': ".2%", 'increasing': {'color': "green"}},
                gauge={
                    'axis': {'range': [None, 1], 'tickwidth': 1, 'tickcolor': "gray", "tickvals": [0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1], "tickformat": "0%"},
                    'bar': {'color': "skyblue"},
                    'bgcolor': "white",
                    'borderwidth': 2,
                    'bordercolor': "gray",
                    'steps': [
                        {'range': [0, 0.4], 'color': 'red'},
                        {'range': [0.4, 0.5], 'color': 'orangered'},
                        {'range': [0.5, 0.6], 'color': 'orange'},
                        {'range': [0.6, 0.7], 'color': 'yellowgreen'},
                        {'range': [0.7, 1], 'color': 'green'}],
                    'threshold': {
                        'line': {'color': "red", 'width': 4},
                        'thickness': 0.75,
                        'value': 0.7}}))

            gauge.update_layout(paper_bgcolor=None, font={'color': "white", 'family': "Arial"})
            st.write(" ")
            st.write(" ")
            st.write(" ")
            st.plotly_chart(gauge)
        except NameError:
            pass
    # topKPI
    st.markdown("---")
    totalshipments = int(df_selection["Waybill Serial Number"].count())
    totalexceeded = int(df_selection.loc[(df_selection["Days of stay"] == "More than 7 Days") | (df_selection["Days of stay"] == "More than 10 Days (Lost?)"), "Waybill Serial Number"].count())
    totalexceeded2 = df_selection.loc[(df_selection["Days of stay"] == "More than 7 Days") | (df_selection["Days of stay"] == "More than 10 Days (Lost?)"), "Waybill Serial Number"].to_list()
    totalexceeded0 = int(df_selection.loc[(df_selection["Days of stay"] == "More than 6 Days (Critical!)"), "Waybill Serial Number"].count())
    totalexceeded02 = df_selection.loc[(df_selection["Days of stay"] == "More than 6 Days (Critical!)"), "Waybill Serial Number"].to_list()
    totalexceeded00 = int(df_selection.loc[(df_selection["Days of stay"] == "More than 7 Days"), "Waybill Serial Number"].count())
    totalexceeded03 = df_selection.loc[(df_selection["Days of stay"] == "More than 7 Days"), "Waybill Serial Number"].to_list()
    totalexceeded000 = int(df_selection.loc[(df_selection["Days of stay"] == "More than 10 Days (Lost?)"), "Waybill Serial Number"].count())
    totalexceeded04 = df_selection.loc[(df_selection["Days of stay"] == "More than 10 Days (Lost?)"), "Waybill Serial Number"].to_list()
    noreturn = int(df_selection.loc[(df_selection["Delay Reason"] == "No Return Application"), "Waybill Serial Number"].count())
    noreturn2 = df_selection.loc[(df_selection["Delay Reason"] == "No Return Application"), "Waybill Serial Number"].to_list()
    noreturn0 = int(df_selection.loc[(df_selection["Delay Reason"] == "No Return Application") & (df_selection["All Abnormal Sequence (Old to New)"].str.contains("customer refuse to take", case=False)), "Waybill Serial Number"].count())
    noreturn02 = df_selection.loc[(df_selection["Delay Reason"] == "No Return Application") & (df_selection["All Abnormal Sequence (Old to New)"].str.contains("customer refuse to take", case=False)), "Waybill Serial Number"].to_list()
    noreturn00 = int(df_selection.loc[(df_selection["Delay Reason"] == "No Return Application") & (df_selection["All Abnormal Sequence (Old to New)"].str.contains("customer refuse to take", case=False) == False), "Waybill Serial Number"].count())
    noreturn002 = df_selection.loc[(df_selection["Delay Reason"] == "No Return Application") & (df_selection["All Abnormal Sequence (Old to New)"].str.contains("customer refuse to take", case=False) == False), "Waybill Serial Number"].to_list()
    unprinted = int(df_selection.loc[(df_selection["Print status"] == "Unprinted") & (df_selection["Application Status"] == "Reviewed"), "Waybill Serial Number"].count())
    unprinted2 = df_selection.loc[(df_selection["Print status"] == "Unprinted") & (df_selection["Application Status"] == "Reviewed"), "Waybill Serial Number"].to_list()
    nodel = int(df_selection.loc[(df_selection["Multiple Deliveries"].isna()), "Waybill Serial Number"].count())
    nodel2 = df_selection.loc[(df_selection["Multiple Deliveries"].isna()), "Waybill Serial Number"].to_list()
    nodel0 = int(df_selection.loc[(df_selection["Multiple Deliveries"].isna()) & (df_selection["Count of Abnormals"].isna()), "Waybill Serial Number"].count())
    nodel02 = df_selection.loc[(df_selection["Multiple Deliveries"].isna()) & (df_selection["Count of Abnormals"].isna()), "Waybill Serial Number"].to_list()
    nodel00 = int(df_selection.loc[(df_selection["Multiple Deliveries"].isna()) & (df_selection["Count of Abnormals"].notna()), "Waybill Serial Number"].count())
    nodel002 = df_selection.loc[(df_selection["Multiple Deliveries"].isna()) & (df_selection["Count of Abnormals"].notna()), "Waybill Serial Number"].to_list()
    noans = int(df_selection.loc[(df_selection['Count of "No Answer" Abnormal'] > 1), "Waybill Serial Number"].count())
    noans2 = df_selection.loc[(df_selection['Count of "No Answer" Abnormal'] > 1), "Waybill Serial Number"].to_list()
    noupdate = int(df_selection.loc[(df_selection['no update'] >= 1), "Waybill Serial Number"].count())
    noupdate2 = df_selection.loc[(df_selection['no update'] >= 1), "Waybill Serial Number"].to_list()
    noupdate01 = int(df_selection.loc[(df_selection['no update'] >= 1) & (df_selection['no update'] < 2), "Waybill Serial Number"].count())
    noupdate12 = df_selection.loc[(df_selection['no update'] >= 1) & (df_selection['no update'] < 2), "Waybill Serial Number"].to_list()
    noupdate02 = int(df_selection.loc[(df_selection['no update'] >= 2) & (df_selection['no update'] < 3), "Waybill Serial Number"].count())
    noupdate22 = df_selection.loc[(df_selection['no update'] >= 2) & (df_selection['no update'] < 3), "Waybill Serial Number"].to_list()
    noupdate03 = int(df_selection.loc[(df_selection['no update'] > 3), "Waybill Serial Number"].count())
    noupdate32 = df_selection.loc[(df_selection['no update'] > 3), "Waybill Serial Number"].to_list()
    ofd3 = int(df_selection.loc[(df_selection['No OFD for more than 3 days'].notna()), "Waybill Serial Number"].count())
    ofd32 = df_selection.loc[(df_selection['No OFD for more than 3 days'].notna()), "Waybill Serial Number"].to_list()
    returnprinted = int(df_selection.loc[(df_selection['printed and not moved'].notna()), "Waybill Serial Number"].count())
    returnprinted2 = df_selection.loc[(df_selection['printed and not moved'].notna()), "Waybill Serial Number"].to_list()
    unprintedold = int(df_selection.loc[(df_selection['unprintedold'].notna()), "Waybill Serial Number"].count())
    unprintedold2 = df_selection.loc[(df_selection['unprintedold'].notna()), "Waybill Serial Number"].to_list()
    unprintedolder = int(df_selection.loc[(df_selection['unprintedolder'].notna()), "Waybill Serial Number"].count())
    unprintedolder2 = df_selection.loc[(df_selection['unprintedolder'].notna()), "Waybill Serial Number"].to_list()
    unprintednew = int(df_selection.loc[(df_selection['unprintednew'].notna()), "Waybill Serial Number"].count())
    unprintednew2 = df_selection.loc[(df_selection['unprintednew'].notna()), "Waybill Serial Number"].to_list()
    autosign = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] >= 10), "Waybill Serial Number"].count())
    autosign2 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] >= 10), "Waybill Serial Number"].to_list()
    autosign10 = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 10), "Waybill Serial Number"].count())
    autosign102 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 10), "Waybill Serial Number"].to_list()
    autosign11 = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 11), "Waybill Serial Number"].count())
    autosign112 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 11), "Waybill Serial Number"].to_list()
    autosign12 = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 12), "Waybill Serial Number"].count())
    autosign122 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 12), "Waybill Serial Number"].to_list()
    autosign13 = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 13), "Waybill Serial Number"].count())
    autosign132 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 13), "Waybill Serial Number"].to_list()
    autosign14 = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 14), "Waybill Serial Number"].count())
    autosign142 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 14), "Waybill Serial Number"].to_list()
    autosign15 = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 15), "Waybill Serial Number"].count())
    autosign152 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] == 15), "Waybill Serial Number"].to_list()
    autosign16 = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] > 15), "Waybill Serial Number"].count())
    autosign162 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] > 15), "Waybill Serial Number"].to_list()

    printed = int(df_selection.loc[(df_selection['Print Time'].notna()), "Waybill Serial Number"].count())
    printednew = int(df_selection.loc[(df_selection['pnmnew'].notna()), "Waybill Serial Number"].count())
    printednew2 = df_selection.loc[(df_selection['pnmnew'].notna()), "Waybill Serial Number"].to_list()
    printedold = int(df_selection.loc[(df_selection['pnmold'].notna()), "Waybill Serial Number"].count())
    printedold2 = df_selection.loc[(df_selection['pnmold'].notna()), "Waybill Serial Number"].to_list()
    printedolder = int(df_selection.loc[(df_selection['pnmolder'].notna()), "Waybill Serial Number"].count())
    printedolder2 = df_selection.loc[(df_selection['pnmolder'].notna()), "Waybill Serial Number"].to_list()

    highpr = int(df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] >= 15), "Waybill Serial Number"].count())
    highpr2 = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] >= 15), "Waybill Serial Number"].to_list()

    unpickupskpi = unpickups
    unpickupskpi2 = unpickupdf_selection

    one, two, three, four, five, six, seven, eight, nine = st.columns(9)
    three0, three1, three2, three3, three00 = st.columns(5)
    a001, a002, a003, a004, a005, a006, a007 = st.columns(7)

    with one:
        st.metric("Exceeded Life-Cycle", f"{totalexceeded:,}", help="شحنات بقالها في الفرع 7 ايام او اكتر")
        if st.button("View", key=2):
            with three1:
                st.metric("More than 6 Days (Critical!)", f"{totalexceeded0:,}")
                st.write("\n".join(totalexceeded02))
            with three2:
                st.metric("More than 7 Days", f"{totalexceeded00:,}")
                st.write("\n".join(totalexceeded03))
            with three3:
                st.metric("More than 10 Days (Lost?)", f"{totalexceeded000:,}")
                st.write("\n".join(totalexceeded04))

    with two:
        st.metric("No Return App", f"{noreturn:,}", help="شحنات المفروض يتعملها ريترن، سواء مرفوضة او بقالها اكتر من 5 ايام في الفرع")
        if st.button("View", key=3):
            with three1:
                st.metric("Customer Refuse to Take", f"{noreturn0:,}")
                st.write("\n".join(noreturn02))
            with three2:
                st.metric("Exceeded 5 Days", f"{noreturn00:,}")
                st.write("\n".join(noreturn002))

    with three:
        st.metric("Return Unprinted", f"{unprinted:,}", help="شحنات اتوافق على الريترن ابلكيشن بتاعها ولسة متطبعتش")
        if st.button("View", key=4):
            with three1:
                st.metric("Approved Today", f"{unprintednew:,}")
                st.write("\n".join(unprintednew2))
            with three2:
                st.metric("Approved Yesterday", f"{unprintedold:,}")
                st.write("\n".join(unprintedold2))
            with three3:
                st.metric("Approved Before That", f"{unprintedolder:,}")
                st.write("\n".join(unprintedolder2))

    with four:
        st.metric('Printed, not Moved', f"{printed:,}", help="شحنات اتطبعلها بوليصة الريترن ولسة مخرجتش من الفرع")
        if st.button("View", key=8):
            with three1:
                st.metric("Printed Today", f"{printednew:,}")
                st.write("\n".join(printednew2))
            with three2:
                st.metric("Printed Yesterday", f"{printedold:,}")
                st.write("\n".join(printedold2))
            with three3:
                st.metric("Printed Before That", f"{printedolder:,}")
                st.write("\n".join(printedolder2))
    with five:
        st.metric("3 Days No OFD", f"{ofd3:,}", help="شحنات بقالها 3 ايام مخرجتش، ومش واخدة لا تأجيل ولا ريترن ابلكيشن")
        if st.button("View", key=5):
            st.write("\n".join(ofd32))

    with six:
        st.metric("No OFD", f"{nodel:,}", help="شحنات مخرجتش ولا محاولة")
        if st.button("View", key=6):
            with three2:
                st.metric("With Abnormal", f"{nodel00:,}")
                st.write("\n".join(nodel002))
            with three3:
                st.metric("Without Abnormal", f"{nodel0:,}")
                st.write("\n".join(nodel02))

    with seven:
        st.metric('No Update', f"{noupdate:,}", help="شحنات مخدتش ابديت بقالها اكتر من 24 ساعة")
        if st.button("View", key=7):
            with three1:
                st.metric("24 Hours", f"{noupdate01:,}")
                st.write("\n".join(noupdate12))
            with three2:
                st.metric("48 Hours", f"{noupdate02:,}")
                st.write("\n".join(noupdate22))
            with three3:
                st.metric("72+ Hours", f"{noupdate03:,}")
                st.write("\n".join(noupdate32))
    with nine:
        st.metric("Unpickups", f"{int(unpickupskpi):,}", help="شحنات المفروض يتعملها بيك اب وبقالها ايام، مترتبة من الاقدم للأحدث")
        if st.button("View", key=10):
            with three00:
                st.table(unpickupdf_selection)
    with eight:
        st.metric("High Priority !!", f"{highpr:,}", help="شحنات قديمة وعمالة تروح وتيجي بين الفروع")
        if st.button("View", key=15):
            st.write("\n".join(highpr2))

    st.markdown("---")
    ### Search bar
    ones, twos, threes, fours = st.columns(4)
    selected = ones.text_input("Waybill Number Search...")
    if selected:
        selected = re.sub('[^a-zA-Z0-9]', ' ', selected).split(" ")
        show = df[["Waybill Serial Number", "Action"]].query('`Waybill Serial Number` == @selected')
        show2 = df.query('`Waybill Serial Number` == @selected')
        st.write("\n:red_circle: Action Needed")
        st.dataframe(show)
        st.write("\n:page_with_curl: Full Data")
        gbb = GridOptionsBuilder.from_dataframe(show2)
        gbb.configure_pagination(paginationAutoPageSize=True)  # Add pagination

        # gb.configure_selection('multiple', use_checkbox=True, groupSelectsChildren="Group checkbox select children") #Enable multi-row selection
        gridOptionss = gbb.build()

        grid_responsess = AgGrid(
            show2,
            gridOptions=gridOptionss,
            data_return_mode='AS_INPUT',
            update_mode='MODEL_CHANGED',
            fit_columns_on_grid_load=False,
            theme='streamlit',  # Add theme color to the table
            enable_enterprise_modules=True,
            height=500,
            width='100%',
            reload_data=False
        )
    st.markdown("---")

    ### TABLES
    tab1, tab2 = st.tabs(["📈 Charts", "🗃 Summary"])

    existing_shipments_br = df_selection.groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    existing_shipments_br.rename(columns={'Waybill Serial Number': 'Existing Shipments'}, inplace=True)

    unsigned_shipments_br = df_selection.loc[(df_selection["Days of stay"] == "More than 7 Days") | (df_selection["Days of stay"] == "More than 10 Days (Lost?)")].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    unsigned_shipments_br.rename(columns={'Waybill Serial Number': 'Exceeded life cycle shipments'}, inplace=True)

    no_return_app_br = df_selection.loc[(df_selection["Delay Reason"] == "No Return Application")].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    no_return_app_br.rename(columns={'Waybill Serial Number': 'No Return Application'}, inplace=True)

    unprinted_shipments_br = df_selection.loc[(df_selection["Print status"] == "Unprinted") & (df_selection["Application Status"] == "Reviewed")].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    unprinted_shipments_br.rename(columns={'Waybill Serial Number': 'Return applications to be printed'}, inplace=True)

    no_ofd_shipments_br = df_selection.loc[(df_selection["Multiple Deliveries"].isna())].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    no_ofd_shipments_br.rename(columns={'Waybill Serial Number': 'Didn\'t get out for delivery'}, inplace=True)

    no_answer_shipments_br = df_selection.loc[(df_selection['Count of "No Answer" Abnormal'] > 1)].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    no_answer_shipments_br.rename(columns={'Waybill Serial Number': 'More than 1 "No Answer" abnormal'}, inplace=True)

    no_update_shipments_br = df_selection.loc[(df_selection['no update'] > 1)].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    no_update_shipments_br.rename(columns={'Waybill Serial Number': 'No Update since More 24 Hours'}, inplace=True)

    ofd3_shipments_br = df_selection.loc[(df_selection['No OFD for more than 3 days'].notna())].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    ofd3_shipments_br.rename(columns={'Waybill Serial Number': 'No Attempts for 3 days'}, inplace=True)

    returnprinted_shipments_br = df_selection.loc[(df_selection['printed and not moved'].notna())].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    returnprinted_shipments_br.rename(columns={'Waybill Serial Number': 'Printed before Today and not Moved'}, inplace=True)

    unprintednew_shipments_br = df_selection.loc[(df_selection['unprintednew'].notna())].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    unprintednew_shipments_br.rename(columns={'Waybill Serial Number': 'Unprinted Approved Today'}, inplace=True)

    unprintedold_shipments_br = df_selection.loc[(df_selection['unprintedold'].notna())].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    unprintedold_shipments_br.rename(columns={'Waybill Serial Number': 'Unprinted Approved Yesterday'}, inplace=True)

    unprintedolder_shipments_br = df_selection.loc[(df_selection['unprintedolder'].notna())].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    unprintedolder_shipments_br.rename(columns={'Waybill Serial Number': 'Unprinted Approved Before That'}, inplace=True)

    highpr_br = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] >= 15)].groupby(["Agency", "Latest Scan Branch"]).count()[["Waybill Serial Number"]]
    highpr_br.rename(columns={'Waybill Serial Number': 'High Priority Shipments'}, inplace=True)

    latest_ontime.rename(columns={'Agency Area Name': 'Agency', "Delivery Branch Name": "Latest Scan Branch"}, inplace=True)
    ontime_br = latest_ontime.groupby(["Agency", "Latest Scan Branch"]).sum()[["On-Time Sign Rate"]]

    funcs = [ontime_br, unsigned_shipments_br, no_return_app_br, unprinted_shipments_br, unprintednew_shipments_br, unprintedold_shipments_br, unprintedolder_shipments_br, returnprinted_shipments_br, ofd3_shipments_br,
             no_ofd_shipments_br, highpr_br, no_answer_shipments_br, no_update_shipments_br]

    final = pd.concat(funcs, axis=1).fillna(0)
    final["On-Time Sign Rate"] = final["On-Time Sign Rate"].round(2).astype(str) + "%"
    final = final.reset_index()

    existing_shipments_ag = df.groupby(["Agency"]).count()[["Waybill Serial Number"]]
    existing_shipments_ag.rename(columns={'Waybill Serial Number': 'Existing Shipments'}, inplace=True)

    unsigned_shipments_ag = df.loc[(df["Days of stay"] == "More than 7 Days") | (df["Days of stay"] == "More than 10 Days (Lost?)")].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    unsigned_shipments_ag.rename(columns={'Waybill Serial Number': 'Exceeded life cycle shipments'}, inplace=True)

    no_return_app_ag = df.loc[(df["Delay Reason"] == "No Return Application")].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    no_return_app_ag.rename(columns={'Waybill Serial Number': 'No Return Application'}, inplace=True)

    unprinted_shipments_ag = df.loc[(df["Print status"] == "Unprinted") & (df["Application Status"] == "Reviewed")].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    unprinted_shipments_ag.rename(columns={'Waybill Serial Number': 'Return applications to be printed'}, inplace=True)

    no_ofd_shipments_ag = df.loc[(df["Multiple Deliveries"].isna())].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    no_ofd_shipments_ag.rename(columns={'Waybill Serial Number': 'Didn\'t get out for delivery'}, inplace=True)

    no_answer_shipments_ag = df.loc[(df['Count of "No Answer" Abnormal'] > 1)].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    no_answer_shipments_ag.rename(columns={'Waybill Serial Number': 'More than 1 "No Answer" abnormal'}, inplace=True)

    no_update_shipments_ag = df.loc[(df['no update'] > 1)].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    no_update_shipments_ag.rename(columns={'Waybill Serial Number': 'No Update since More 24 Hours'}, inplace=True)

    ofd3_shipments_ag = df.loc[(df['No OFD for more than 3 days'].notna())].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    ofd3_shipments_ag.rename(columns={'Waybill Serial Number': 'No Attempts for 3 days'}, inplace=True)

    returnprinted_shipments_ag = df.loc[(df['printed and not moved'].notna())].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    returnprinted_shipments_ag.rename(columns={'Waybill Serial Number': 'Printed before Today and not Moved'}, inplace=True)

    unprintednew_shipments_ag = df_selection.loc[(df_selection['unprintednew'].notna())].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    unprintednew_shipments_ag.rename(columns={'Waybill Serial Number': 'Unprinted Approved Today'}, inplace=True)

    unprintedold_shipments_ag = df_selection.loc[(df_selection['unprintedold'].notna())].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    unprintedold_shipments_ag.rename(columns={'Waybill Serial Number': 'Unprinted Approved Yesterday'}, inplace=True)

    unprintedolder_shipments_ag = df_selection.loc[(df_selection['unprintedolder'].notna())].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    unprintedolder_shipments_ag.rename(columns={'Waybill Serial Number': 'Unprinted Approved Before That'}, inplace=True)

    highpr_ag = df_selection.loc[(df_selection['Auto Sign Shipments (Critical!)'] >= 15)].groupby(["Agency"]).count()[["Waybill Serial Number"]]
    highpr_ag.rename(columns={'Waybill Serial Number': 'High Priority Shipments'}, inplace=True)

    funcs2 = [unsigned_shipments_ag, no_return_app_ag, unprinted_shipments_ag, unprintednew_shipments_ag, unprintedold_shipments_ag, unprintedolder_shipments_ag, returnprinted_shipments_ag, ofd3_shipments_ag,
              no_ofd_shipments_ag, highpr_ag, no_answer_shipments_ag, no_update_shipments_ag]

    final2 = pd.concat(funcs2, axis=1).fillna(0)
    final2[final2.columns] = final2[final2.columns].astype(int)
    final2 = final2.reset_index()
    with tab2:
        gb = GridOptionsBuilder.from_dataframe(final2)
        gb.configure_pagination(paginationAutoPageSize=True)  # Add pagination

        # gb.configure_selection('multiple', use_checkbox=True, groupSelectsChildren="Group checkbox select children") #Enable multi-row selection
        gridOptions = gb.build()

        grid_response = AgGrid(
            final2,
            gridOptions=gridOptions,
            data_return_mode='AS_INPUT',
            update_mode='MODEL_CHANGED',
            fit_columns_on_grid_load=False,
            theme='streamlit',  # Add theme color to the table
            enable_enterprise_modules=True,
            height=230,
            width='100%',
            reload_data=False
        )
        st.markdown("##")
        gb2 = GridOptionsBuilder.from_dataframe(final)
        gb2.configure_pagination(paginationAutoPageSize=True)  # Add pagination
        # gb.configure_selection('multiple', use_checkbox=True, groupSelectsChildren="Group checkbox select children") #Enable multi-row selection
        gridOptions2 = gb2.build()

        grid_response2 = AgGrid(
            final,
            gridOptions=gridOptions2,
            data_return_mode='AS_INPUT',
            update_mode='MODEL_CHANGED',
            fit_columns_on_grid_load=False,
            theme='streamlit',  # Add theme color to the table
            enable_enterprise_modules=True,
            height=850,
            width='100%',
            reload_data=False
        )

    ### PLOTS
    with tab1:
        import plotly.graph_objects as go

        # Define the target on-time rate and the color palette
        target_rate = 70
        color_scale = [[0, 'black'], [0.4, 'darkred'], [0.6, 'red'], [0.65, 'yellow'], [0.7, 'green'], [0.85, 'darkgreen'], [1, 'darkgreen']]


        # Prepare on-time data
        def calculate_ontime_sign_rate(ontimedf):
            grouped = ontimedf.groupby(["Date", "Agency Area Name", "Delivery Branch Name"])[['Receivable Amount', 'On-time signing Amount']].sum().reset_index()
            latest_date = grouped["Date"].max()
            latest_grouped = grouped.query(f"Date == '{latest_date}'")
            renamed = latest_grouped.rename(columns={'Agency Area Name': 'Agency', "Delivery Branch Name": "Latest Scan Branch"})
            ontime_chart = renamed.groupby(["Agency", "Latest Scan Branch"]).sum()[["On-time signing Amount", "Receivable Amount"]]
            ontime_chart["On-Time Sign Rate"] = ((ontime_chart["On-time signing Amount"] / ontime_chart['Receivable Amount']) * 100).round(2)
            return ontime_chart[["On-Time Sign Rate"]]


        ontime_chart = calculate_ontime_sign_rate(ontimedf)

        # Create the bar chart
        ontime_chart = (
            ontime_chart
            .reset_index()
            .drop('Agency', axis=1)
            .rename(columns={"Latest Scan Branch": "Branch Name"})
            .sort_values(by="On-Time Sign Rate", ascending=True)
        )
        ontimefig = px.bar(
            ontime_chart,
            y='Branch Name',
            x='On-Time Sign Rate',
            text=[f'{i}%' for i in ontime_chart['On-Time Sign Rate']],
            color='On-Time Sign Rate',
            color_continuous_scale=color_scale,
            range_color=[0, 100],
            orientation='h',
            height=600,
            width=800,
            hover_data={'On-Time Sign Rate': ':.2f'}
        )

        # Update the chart layout
        ontimefig.update_traces(marker_line_color='grey', marker_line_width=1)
        ontimefig.update_layout(
            title="On-Time Sign Rate by Branch",
            title_font_size=24,
            title_font_family="Arial",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(size=16),
            xaxis=dict(showgrid=False, title="On-Time Sign Rate (%)", title_font=dict(size=18)),
            yaxis=dict(showgrid=False, title="Branch", title_font=dict(size=18), categoryorder='total ascending', range=[-0.5, len(ontime_chart) - 0.5]),
            margin=dict(l=0, r=0, t=60, b=0),
            annotations=[
                dict(
                    text=f"On-time rate should reach {target_rate}%. On-time rate below 60% is considered very poor performance.",
                    x=0.5,
                    y=1.08,
                    xref='paper',
                    yref='paper',
                    showarrow=False,
                    font=dict(size=14, color='grey')
                )
            ]
        )

        # Show the actual values next to the bars
        ontimefig.update_traces(textposition='outside')

        # Use a consistent y-axis scale
        ontimefig.update_yaxes(scaleanchor='x', scaleratio=1)

        # Add a dashed red line for the target on-time rate
        ontimefig.add_shape(
            type='line',
            x0=-0.5,
            x1=len(ontime_chart) - 0.5,
            y0=70,
            y1=70,
            line=dict(color='red', width=2, dash='dash')
        )

        # Display the chart
        st.plotly_chart(ontimefig, use_container_width=True)

        ###

        plot1df = df.query('Agency == @agency')
        noofd = plot1df.loc[(df["Multiple Deliveries"].isna())].groupby(by='Latest Scan Branch').count()[["Waybill Serial Number"]].sort_values(by="Waybill Serial Number")

        fig_noofd = go.Figure(data=[go.Bar(
            x=noofd["Waybill Serial Number"],
            y=noofd.index,
            orientation='h',
            marker=dict(color=noofd["Waybill Serial Number"], coloraxis="coloraxis"),
            text=noofd["Waybill Serial Number"],
            texttemplate='%{text:.0f}',
            textposition='inside',
            hovertemplate='Branch: %{y}<br>Number of Deliveries: %{x:.0f}<extra></extra>'
        )])

        fig_noofd.update_layout(
            title="Shipments That Didn't Go Out for Delivery by Branch",
            title_font_size=24,
            title_font_family="Arial",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(size=16),
            height=(40 * df.loc[(df["Multiple Deliveries"].isna())].groupby(by='Latest Scan Branch').count().shape[0]) + 150,
            xaxis=dict(showgrid=False, title="Number of Deliveries", title_font=dict(size=18)),
            yaxis=dict(showgrid=False, title="Branch", title_font=dict(size=18)),
            coloraxis=dict(colorscale='Reds'),
            margin=dict(l=0, r=0, t=60, b=0)
        )

        fig_noofd.update_traces(marker_line_width=0)

        fig_noofd.add_layout_image(
            dict(
                source="https://i.imgur.com/XXXXXX.png",  # Replace with your image source
                xref="paper",
                yref="paper",
                x=1.1,
                y=1.05,
                sizex=0.2,
                sizey=0.2,
                xanchor="right",
                yanchor="bottom"
            )
        )

        fig_noofd.update_layout(hovermode="closest")

        st.plotly_chart(fig_noofd, use_container_width=True)

        # Filter the dataframe by the selected agency
        plot2df = df.query('Agency == @agency')

        # Group the data by agency and latest scan branch and count the number of shipments in each life cycle category
        try:
            allshipments2 = plot2df.groupby(by=["Agency", "Latest Scan Branch"])['exceeded'].value_counts().unstack().fillna(0).reset_index().sort_values(by="Exceeded", ascending=False)
        except:
            try:
                allshipments2 = plot2df.groupby(by=["Agency", "Latest Scan Branch"])['exceeded'].value_counts().unstack().fillna(0).reset_index().sort_values(by="About to Exceed", ascending=False)
            except:
                allshipments2 = plot2df.groupby(by=["Agency", "Latest Scan Branch"])['exceeded'].value_counts().unstack().fillna(0).reset_index().sort_values(by="Not Exceed", ascending=False)

        # Melt the data to create a stacked bar chart
        stack = allshipments2.melt(id_vars=['Agency', 'Latest Scan Branch'], value_vars=['About to Exceed', 'Exceeded'], var_name='Life Cycle')

        # Set the colors for each life cycle category
        rating_color = ['#ffff00', '#ff0000']

        # Create a new plotly figure
        fig = go.Figure()

        # Add each life cycle category to the figure as a separate trace
        for r, c in zip(stack['Life Cycle'].unique(), rating_color):
            stack_plot = stack[stack['Life Cycle'] == r]
            fig.add_trace(
                go.Bar(
                    x=stack_plot['Latest Scan Branch'],
                    y=stack_plot['value'],
                    marker_color=c,
                    name=r,
                    text=stack_plot['value'],
                    textposition='auto'
                )
            )

        # Customize the layout of the figure
        fig.update_layout(
            title="<b>Shipment Life Cycles by Latest Scan Branch</b>",
            title_font_size=24,
            title_font_family="Arial",
            font=dict(size=16),
            xaxis_title="Latest Scan Branch",
            yaxis_title="Number of Shipments",
            plot_bgcolor='rgba(0,0,0,0)',
            barmode='stack',
            legend=dict(
                x=0.7,
                y=1,
                bgcolor='grey',
                bordercolor="black",
                borderwidth=1
            ),
            margin=dict(l=10, r=10, t=100, b=10),
            height=600
        )

        # Show the figure
        left, right = st.columns(2)
        left.plotly_chart(fig, use_container_width=True)

        # Pie Chart
        stackpie = stack[(stack['Life Cycle'] == "Exceeded") & (stack['value'] > 0)]
        total_exceeded = stackpie['value'].sum()
        stackpie['Percentage'] = 100 * stackpie['value'] / total_exceeded
        fig_pie = px.pie(stackpie, values="Percentage", names='Latest Scan Branch', title="<b>Exceeded Life Cycle Shipments Per Branch</b>",
                         color_discrete_sequence=px.colors.qualitative.Dark2,
                         labels={'Percentage': 'Percentage of Exceeded Life Cycle Shipments'},
                         hole=0.5, )
        fig_pie.update_layout(title_font_size=24,
                              title_font_family="Arial",
                              font=dict(size=16), )
        right.plotly_chart(fig_pie, use_container_width=True)
    st.markdown("---")

with tabb2:
    outcol, arrcol, signcol, abncol = st.columns(4)
    with signcol:
        signscan = st.file_uploader("Signing Scans File")
        if signscan is not None:
            signscandf = pd.read_excel(signscan)
    with outcol:
        ofdscan = st.file_uploader("Out for Delivery Scans File")
        if ofdscan is not None:
            ofdscandf = pd.read_excel(ofdscan)
    with arrcol:
        arrscan = st.file_uploader("Arrival Scans File")
        if arrscan is not None:
            arrscandf = pd.read_excel(arrscan)
    with abncol:
        abnscan = st.file_uploader("Abnormal Registrations File")
        if abnscan is not None:
            abnscandf = pd.read_excel(abnscan)
    courieractions, couriermap = st.tabs(["Couriers Actions", "Activity Map"])
    with courieractions:
        sorting, first_sign = st.columns([1, 3])
        with sorting:
            if ofdscan is not None and arrscan is not None:
                ofdscandf = ofdscandf.append(arrscandf, ignore_index=True)
                ofdscandf[[col for col in ofdscandf.columns if 'time' in col or 'Time' in col]] = ofdscandf[[col for col in ofdscandf.columns if 'time' in col or 'Time' in col]].apply(pd.to_datetime)
                st.write("First Arrival Scan: ", ofdscandf.loc[ofdscandf["Scan Type"] == "Arrival Scan", ["Scan time"]].min().to_list()[0])
                st.write("Last Arrival Scan: ", ofdscandf.loc[ofdscandf["Scan Type"] == "Arrival Scan", ["Scan time"]].max().to_list()[0])
                st.write("First Out for Delivery Scan: ", ofdscandf.loc[ofdscandf["Scan Type"] == "Out for Delivery Scan", ["Scan time"]].min().to_list()[0])
                st.write("Last Out for Delivery Scan: ", ofdscandf.loc[ofdscandf["Scan Type"] == "Out for Delivery Scan", ["Scan time"]].max().to_list()[0])
                st.info("Truck Unloading Period: " + str((ofdscandf.loc[ofdscandf["Scan Type"] == "Arrival Scan", ["Scan time"]].max().to_list()[0] - ofdscandf.loc[ofdscandf["Scan Type"] == "Arrival Scan", ["Scan time"]].min()).to_list()[0]))

                if ofdscandf.loc[ofdscandf["Scan Type"] == "Out for Delivery Scan", ["Scan time"]].min().to_list()[0] > ofdscandf.loc[ofdscandf["Scan Type"] == "Arrival Scan", ["Scan time"]].max().to_list()[0]:
                    st.info("Sorting Period: " + str((ofdscandf.loc[ofdscandf["Scan Type"] == "Out for Delivery Scan", ["Scan time"]].max().to_list()[0] - ofdscandf.loc[ofdscandf["Scan Type"] == "Arrival Scan", ["Scan time"]].max()).to_list()[0]))
                else:
                    st.info(
                        "Sorting Period: " + str(
                            (ofdscandf.loc[ofdscandf["Scan Type"] == "Out for Delivery Scan", ["Scan time"]].max().to_list()[0] - ofdscandf.loc[ofdscandf["Scan Type"] == "Out for Delivery Scan", ["Scan time"]].min()).to_list()[0]))
                st.info("Sorted Shipments: " + str(ofdscandf.loc[:, "Waybill NO."].nunique()))

        with first_sign:
            if signscan is not None and abnscan is None and ofdscan is None:
                signscandf[[col for col in signscandf.columns if 'time' in col or 'Time' in col]] = signscandf[[col for col in signscandf.columns if 'time' in col or 'Time' in col]].apply(pd.to_datetime)
                signscandf["Scan time"] = pd.Series([val.time() for val in signscandf["Scan time"]])
                result = signscandf.groupby('Delivery or pickup Courier')["Scan time"].agg('min').reset_index().sort_values(by="Scan time", ascending=False).reset_index().rename(columns={"Scan time": "First Sign"}).drop(["index"], axis=1)

                gbsign = GridOptionsBuilder.from_dataframe(result)
                gbsign.configure_pagination(paginationAutoPageSize=True)  # Add pagination
                # gb.configure_selection('multiple', use_checkbox=True, groupSelectsChildren="Group checkbox select children") #Enable multi-row selection
                gridOptionssign = gbsign.build()
                grid_responsesign = AgGrid(
                    result,
                    gridOptions=gridOptionssign,
                    data_return_mode='AS_INPUT',
                    update_mode='MODEL_CHANGED',
                    fit_columns_on_grid_load=True,
                    theme='streamlit',  # Add theme color to the table
                    enable_enterprise_modules=True,
                    height=1200,
                    width='100%',
                    reload_data=False
                )
            elif signscan is not None and abnscan is None and ofdscan is not None:
                signscandf[[col for col in signscandf.columns if 'time' in col or 'Time' in col]] = signscandf[[col for col in signscandf.columns if 'time' in col or 'Time' in col]].apply(pd.to_datetime)
                signscandf["Scan time"] = pd.Series([val.time() for val in signscandf["Scan time"]])
                result = signscandf.groupby('Delivery or pickup Courier')["Scan time"].agg('min').reset_index().sort_values(by="Scan time", ascending=False).reset_index().rename(columns={"Scan time": "First Sign"}).drop(["index"], axis=1)
                result = result.merge(ofdscandf[["Delivery or pickup Courier"]].drop_duplicates(), how="outer")
                result = result.dropna(how="all")
                gbsign = GridOptionsBuilder.from_dataframe(result)
                gbsign.configure_pagination(paginationAutoPageSize=True)  # Add pagination
                # gb.configure_selection('multiple', use_checkbox=True, groupSelectsChildren="Group checkbox select children") #Enable multi-row selection
                gridOptionssign = gbsign.build()
                grid_responsesign = AgGrid(
                    result,
                    gridOptions=gridOptionssign,
                    data_return_mode='AS_INPUT',
                    update_mode='MODEL_CHANGED',
                    fit_columns_on_grid_load=True,
                    theme='streamlit',  # Add theme color to the table
                    enable_enterprise_modules=True,
                    height=1200,
                    width='100%',
                    reload_data=False
                )
            elif signscan is not None and abnscan is not None and ofdscan is not None:
                signscandf = signscandf.append(abnscandf, ignore_index=True)
                signscandf.loc[signscandf["Scan Type"] == "Abnormal parcels scan", "Delivery or pickup Courier"] = signscandf.loc[signscandf["Scan Type"] == "Abnormal parcels scan", "Operator"]
                signscandf[[col for col in signscandf.columns if 'time' in col or 'Time' in col]] = signscandf[[col for col in signscandf.columns if 'time' in col or 'Time' in col]].apply(pd.to_datetime)
                signscandf["Scan time2"] = pd.Series([val.time() for val in signscandf["Scan time"]])
                fs = signscandf.loc[signscandf["Scan Type"] == "Signing scan"].groupby('Delivery or pickup Courier')["Scan time2"].agg('min').reset_index().sort_values(by="Scan time2", ascending=False).reset_index().rename(
                    columns={"Scan time2": "First Sign"}).drop(["index"], axis=1)
                fs = fs.merge(ofdscandf[["Delivery or pickup Courier"]].drop_duplicates(), how="outer")
                last_activity = signscandf.groupby('Delivery or pickup Courier')["Scan time"].agg('max').reset_index().sort_values(by="Scan time", ascending=False).reset_index().rename(columns={"Scan time": "Last Activity Time"}).drop(["index"],
                                                                                                                                                                                                                                           axis=1)
                fs = fs.merge(last_activity, how="left")
                fs["Inactive For .. (Minutes)"] = round(((datetime.datetime.now() - fs["Last Activity Time"]) / np.timedelta64(1, 'm')) + 120, )
                fs["Scan time"] = fs["Last Activity Time"]
                fs = fs.merge(signscandf[["Delivery or pickup Courier", "Scan time", "Waybill NO.", "Scan Type", "Branch latitude and longitude"]], on=["Delivery or pickup Courier", "Scan time"], how="left").drop_duplicates(
                    subset="Delivery or pickup Courier")
                fs[['lon', 'lat']] = fs['Branch latitude and longitude'].str.split(',', expand=True)
                fs["Last Location (GPS)"] = fs["lat"] + "," + fs["lon"]
                fs["Last Activity Type"] = fs["Waybill NO."] + " (" + fs["Scan Type"] + ")"

                fs = fs[["Delivery or pickup Courier", "First Sign", "Last Activity Time", "Last Activity Type", "Last Location (GPS)", "Inactive For .. (Minutes)"]]
                fs = fs.dropna(how='all')

                gbsign = GridOptionsBuilder.from_dataframe(fs)
                gbsign.configure_pagination(paginationAutoPageSize=True)
                gridOptionssign = gbsign.build()
                grid_responsesign = AgGrid(
                    fs,
                    gridOptions=gridOptionssign,
                    data_return_mode='AS_INPUT',
                    update_mode='MODEL_CHANGED',
                    fit_columns_on_grid_load=True,
                    theme='streamlit',  # Add theme color to the table
                    enable_enterprise_modules=True,
                    height=1200,
                    width='100%',
                    reload_data=False
                )

            elif signscan is not None and abnscan is not None and ofdscan is None:
                signscandf = signscandf.append(abnscandf, ignore_index=True)
                signscandf.loc[signscandf["Scan Type"] == "Abnormal parcels scan", "Delivery or pickup Courier"] = signscandf.loc[signscandf["Scan Type"] == "Abnormal parcels scan", "Operator"]
                signscandf[[col for col in signscandf.columns if 'time' in col or 'Time' in col]] = signscandf[[col for col in signscandf.columns if 'time' in col or 'Time' in col]].apply(pd.to_datetime)
                signscandf["Scan time2"] = pd.Series([val.time() for val in signscandf["Scan time"]])
                fs = signscandf.loc[signscandf["Scan Type"] == "Signing scan"].groupby('Delivery or pickup Courier')["Scan time2"].agg('min').reset_index().sort_values(by="Scan time2", ascending=False).reset_index().rename(
                    columns={"Scan time2": "First Sign"}).drop(["index"], axis=1)
                last_activity = signscandf.groupby('Delivery or pickup Courier')["Scan time"].agg('max').reset_index().sort_values(by="Scan time", ascending=False).reset_index().rename(columns={"Scan time": "Last Activity Time"}).drop(["index"],
                                                                                                                                                                                                                                           axis=1)
                fs = fs.merge(last_activity, how="left")
                fs["Inactive For .. (Minutes)"] = round(((datetime.datetime.now() - fs["Last Activity Time"]) / np.timedelta64(1, 'm')) + 120, )
                fs["Scan time"] = fs["Last Activity Time"]
                fs = fs.merge(signscandf[["Delivery or pickup Courier", "Scan time", "Waybill NO.", "Scan Type", "Branch latitude and longitude"]], on=["Delivery or pickup Courier", "Scan time"], how="left").drop_duplicates(
                    subset="Delivery or pickup Courier")
                fs[['lon', 'lat']] = fs['Branch latitude and longitude'].str.split(',', expand=True)
                fs["Last Location (GPS)"] = fs["lat"] + "," + fs["lon"]
                fs["Last Activity Type"] = fs["Waybill NO."] + " (" + fs["Scan Type"] + ")"

                fs = fs[["Delivery or pickup Courier", "First Sign", "Last Activity Time", "Last Activity Type", "Last Location (GPS)", "Inactive For .. (Minutes)"]]
                gbsign = GridOptionsBuilder.from_dataframe(fs)
                gbsign.configure_pagination(paginationAutoPageSize=True)
                gridOptionssign = gbsign.build()
                grid_responsesign = AgGrid(
                    fs,
                    gridOptions=gridOptionssign,
                    data_return_mode='AS_INPUT',
                    update_mode='MODEL_CHANGED',
                    fit_columns_on_grid_load=True,
                    theme='streamlit',  # Add theme color to the table
                    enable_enterprise_modules=True,
                    height=1200,
                    width='100%',
                    reload_data=False
                )
    with couriermap:
        if signscan is not None and abnscan is not None:
            activity = pd.read_excel(signscan).append(pd.read_excel(abnscan), ignore_index=True)
            activity.loc[activity["Scan Type"] == "Abnormal parcels scan", "Delivery or pickup Courier"] = activity.loc[activity["Scan Type"] == "Abnormal parcels scan", "Operator"]
            activity[['lon', 'lat']] = activity['Branch latitude and longitude'].str.split(',', expand=True).astype(float)
            activity["lonlatGPS"] = activity[["lon", "lat"]].values.tolist()
            courierr = st.selectbox("Select Courier:", options=activity["Delivery or pickup Courier"].unique())
            activity_selection = activity.query('`Delivery or pickup Courier` == @courierr').sort_values(by="Scan time")
            path = activity_selection[["Delivery or pickup Courier", "lonlatGPS"]].groupby('Delivery or pickup Courier', as_index=False).agg({'lonlatGPS': list})

            min_lat = activity_selection['lat'].min()
            max_lat = activity_selection['lat'].max()
            min_lon = activity_selection['lon'].min()
            max_lon = activity_selection['lon'].max()
            center_lat = (max_lat + min_lat) / 2.0
            center_lon = (max_lon + min_lon) / 2.0
            range_lon = abs(max_lon - min_lon)
            range_lat = abs(max_lat - min_lat)
            if range_lon > range_lat:
                longitude_distance = range_lon
            else:
                longitude_distance = range_lat


            def _get_zoom_level(distance):
                _ZOOM_LEVELS = [360, 180, 90, 45, 22.5, 11.25, 5.625, 2.813, 1.406, 0.703, 0.352, 0.176, 0.088, 0.044, 0.022, 0.011, 0.005, 0.003, 0.001, 0.0005, 0.00025, ]
                if distance < _ZOOM_LEVELS[-1]:
                    return 12
                for i in range(len(_ZOOM_LEVELS) - 1):
                    if _ZOOM_LEVELS[i + 1] < distance <= _ZOOM_LEVELS[i]:
                        return i


            zoom = _get_zoom_level(longitude_distance)
            if st.checkbox("Show/Hide Data", True):
                st.subheader("Full Activity of " + courierr)

                GGBB = GridOptionsBuilder.from_dataframe(activity_selection.reset_index().drop(["index"], axis=1))
                GGBB.configure_selection('multiple', use_checkbox=True, groupSelectsChildren="Group checkbox select children")  # Enable multi-row selection
                gridOptionsssss = GGBB.build()

                grid_responseeeeeeee = AgGrid(
                    activity_selection.reset_index().drop(["index"], axis=1),
                    gridOptions=gridOptionsssss,
                    data_return_mode='AS_INPUT',
                    update_mode='MODEL_CHANGED',
                    fit_columns_on_grid_load=False,
                    theme='streamlit',  # Add theme color to the table
                    enable_enterprise_modules=True,
                    height=350,
                    width='100%',
                    reload_data=False
                )

                data = grid_responseeeeeeee['data']
                selected = grid_responseeeeeeee['selected_rows']
                dffffff = pd.DataFrame(selected)  # Pass the selected rows to a new dataframe df
            if len(selected) > 0:
                st.pydeck_chart(pdk.Deck(
                    map_style=None,
                    initial_view_state=pdk.ViewState(
                        latitude=center_lat,
                        longitude=center_lon,
                        zoom=zoom,
                    ),
                    tooltip={
                        "text": "{Waybill NO.}\n{Scan Type} ({Abnormal parcel type})\n{Scan time}"
                    },
                    layers=[
                        pdk.Layer(
                            'ScatterplotLayer',
                            dffffff[["Waybill NO.", "Scan time", "Scan Type", "Abnormal parcel type", "Delivery or pickup Courier", "lat", "lon"]].fillna("NaN"),
                            get_position=['lon', 'lat'],
                            auto_highlight=True,
                            get_radius=100,
                            get_fill_color=[255, 'lon > 0 ? 200 * lon : -200 * lon', 'lon', 140],
                            pickable=True)
                    ],
                ))
            else:
                st.pydeck_chart(pdk.Deck(
                    map_style=None,
                    initial_view_state=pdk.ViewState(
                        latitude=center_lat,
                        longitude=center_lon,
                        zoom=zoom,
                    ),
                    tooltip={
                        "text": "{Waybill NO.}\n{Scan Type} ({Abnormal parcel type})\n{Scan time}"
                    },
                    layers=[
                        pdk.Layer(
                            'ScatterplotLayer',
                            activity_selection[["Waybill NO.", "Scan time", "Scan Type", "Abnormal parcel type", "Delivery or pickup Courier", "lat", "lon"]].fillna("NaN"),
                            get_position=['lon', 'lat'],
                            auto_highlight=True,
                            get_radius=100,
                            get_fill_color=[255, 'lon > 0 ? 200 * lon : -200 * lon', 'lon', 140],
                            pickable=True)
                    ],
                ))
