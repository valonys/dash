import streamlit as st
import plotly.graph_objects as go
import pandas as pd
from data_processing import DataProcessor

def apply_delay_colors(val, delay_col):
    delay_colors = {
        '< 6 Months': 'background-color: #90EE90',
        '6 Months < x <1 Yrs': 'background-color: #FFB366',
        '1 Yrs < x <2 Yrs': 'background-color: #FFE5B4',
        '2 Yrs < x <3 Yrs': 'background-color: #FFC0CB',
        '> 3 Yrs': 'background-color: #FF6B6B'
    }
    return delay_colors.get(delay_col, '')

def create_monthly_performance_chart(monthly_data, title="Monthly Performance Overview"):
    months = [item['month'] for item in monthly_data]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        name='Plan',
        x=months,
        y=[item['total_planned'] for item in monthly_data],
        marker_color='rgb(30, 144, 255)',
        text=[item['total_planned'] for item in monthly_data],
        textposition='inside',
        insidetextanchor='middle',
        textfont=dict(size=10, family="Tw Cen MT", color="white")
    ))
    fig.add_trace(go.Bar(
        name='Backlog',
        x=months,
        y=[item['backlog_count'] for item in monthly_data],
        marker_color='rgb(255, 0, 0)',
        text=[item['backlog_count'] for item in monthly_data],
        textposition='inside',
        insidetextanchor='middle',
        textfont=dict(size=10, family="Tw Cen MT", color="white")
    ))
    fig.add_trace(go.Bar(
        name='Completed',
        x=months,
        y=[item['completed_count'] for item in monthly_data],
        marker_color='rgb(0, 255, 0)',
        text=[item['completed_count'] for item in monthly_data],
        textposition='inside',
        insidetextanchor='middle',
        textfont=dict(size=10, family="Tw Cen MT", color="black")
    ))
    fig.add_trace(go.Scatter(
        name='Progress %',
        x=months,
        y=[item['progress_percentage'] for item in monthly_data],
        yaxis='y2',
        line=dict(color='black', width=1),
        text=[f"{item['progress_percentage']}%" for item in monthly_data],
        textposition='top center',
        textfont=dict(size=10, family="Tw Cen MT", color="black")
    ))
    fig.update_layout(
        title=dict(text=title, font=dict(family='Tw Cen MT', size=18), x=0.05, y=0.95),
        barmode='stack',
        yaxis=dict(title='Work Orders', range=[0, None], gridcolor=None,
                  titlefont=dict(size=16, family="Tw Cen MT", color="black")),
        yaxis2=dict(title='Progress %', overlaying='y', side='right', range=[0, 100], gridcolor=None,
                   titlefont=dict(size=16, family="Tw Cen MT", color="black")),
        height=400,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        plot_bgcolor='white'
    )
    return fig

def create_completion_bar_chart(df, processor):
    # Get completion data for all months and aggregate by item class
    item_class_totals = {}
    for month in range(1, 13):  # Assuming 12 months
        monthly_data = processor.get_monthly_item_class_performance(df, month)
        for index, row in monthly_data.iterrows():
            item_class = index  # Assuming index is the item class
            completed = row['Progress Accomplished']
            if item_class in item_class_totals:
                item_class_totals[item_class] += completed
            else:
                item_class_totals[item_class] = completed
    
    # Convert to lists for plotting
    item_classes = list(item_class_totals.keys())
    completed = list(item_class_totals.values())
    
    # Define colors based on completed count
    colors = []
    for value in completed:
        if value > 15:
            colors.append('rgb(0, 128, 0)')  # Green
        elif 5 <= value <= 15:
            colors.append('rgb(255, 165, 0)')  # Orange
        else:
            colors.append('rgb(255, 0, 0)')  # Red
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=item_classes,
        y=completed,
        name='Completed',
        marker_color=colors,
        text=completed,
        textposition='auto',
        textfont=dict(size=14, family="Tw Cen MT", color="black"),
        width=0.7
    ))
    fig.update_layout(
        title=dict(text="Completed Jobs by Item Class", font=dict(family='Tw Cen MT', size=18), x=0.05, y=0.95),
        yaxis=dict(title='(EXDO+QCAP) - Orders', titlefont=dict(size=14, family="Tw Cen MT", color="black")),
        xaxis=dict(title='Item Class', titlefont=dict(size=14, family="Tw Cen MT", color="black")),
        bargap=0.2, 
        height=600,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        plot_bgcolor='white',
        font=dict(family="Tw Cen MT", size=12, color="black")
    )
    return fig

def create_backlog_item_class_chart(df_backlog):
    df_chart = df_backlog.drop(index='Grand Total', columns='Totals')
    totals = df_chart.sum(axis=1)
    
    colors = []
    for value in totals:
        if value > 15:
            colors.append('rgb(255, 0, 0)')
        elif 10 <= value <= 15:
            colors.append('rgb(255, 165, 0)')
        elif 5 <= value < 10:
            colors.append('rgb(255, 215, 0)')
        else:
            colors.append('rgb(0, 128, 0)')
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=df_chart.index,
        x=totals,
        orientation='h',
        marker_color=colors,
        text=totals,
        textposition='auto',
        textfont=dict(size=12, family="Tw Cen MT", color="black"),
        width=0.7,
    ))
    fig.update_layout(
        title=dict(text="Backlog by Item Class", font=dict(family='Tw Cen MT', size=18), x=0.05, y=0.95),
        xaxis=dict(title='Number of Items', titlefont=dict(size=16, family="Tw Cen MT", color="black")),
        yaxis=dict(title='Item Class', titlefont=dict(size=16, family="Tw Cen MT", color="black")),
        bargap=0.1,
        height=600,
        plot_bgcolor='white',
        font=dict(family="Tw Cen MT", size=12)
    )
    return fig

def main():
    st.set_page_config(layout='wide', page_title='Inspection Program Dashboard')
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Tw+Cen+MT:wght@400;700&display=swap');
        .main-title {
            font-family: 'Tw Cen MT';
            font-size: 32px;
            font-weight: bold;
            text-align: center;
            padding: 20px 0;
        }
        .metric-label {
            font-family: 'Tw Cen MT';
            font-size: 18px !important;
        </style>
        """, unsafe_allow_html=True)

    processor = DataProcessor()
    with st.sidebar:
        st.title("Settings")
        selected_site = st.radio("Select Site", ["GIR", "DAL", "PAZ", "CLV"])
        uploaded_file = st.file_uploader("Upload Excel File", type=["xlsm"])

    if uploaded_file and selected_site:
        df = processor.load_inspection_data(uploaded_file, selected_site)
        results = processor.analyze_data(df)
        st.markdown(
            f'<p class="main-title" style="font-size: 32px; color: black; font-family: \'Tw Cen MT\', sans-serif;">Inspection Dashboard - {selected_site}</p>',
            unsafe_allow_html=True)

        st.markdown("""
            <style>
            .stMetric label {
                font-family: 'Tw Cen MT', sans-serif !important;
                font-size: 14px !important;
                font-weight: bold !important;
            }
            .stMetric .css-1wivap2 {
                font-family: 'Tw Cen MT', sans-serif !important;
                font-size: 24px !important;
                font-weight: bold !important;
                color: #000000 !important;
            }
            </style>
            """, unsafe_allow_html=True)

        col1, col2, col3, col4 = st.columns(4)
        with col1:
            backlog = results['backlog_summary']['total_backlog']
            st.markdown(
                f"""
                <div style="font-family: 'Tw Cen MT'; font-size: 18px; font-weight: bold; color: red;">
                    Total Backlog
                </div>
                <div style="font-family: 'Tw Cen MT'; font-size: 18px; font-weight: bold; color: red;">
                    {backlog}
                </div>
                """, unsafe_allow_html=True)
            
        with col2:
            sce_backlog_rate = results['backlog_summary']['sece_metrics']['sece_percentage']
            st.markdown(
                f"""
                <div style="font-family: 'Tw Cen MT'; font-size: 18px; font-weight: bold; color: red;">
                    SCE Backlog Rate
                </div>
                <div style="font-family: 'Tw Cen MT'; font-size: 18px; font-weight: bold; color: red;">
                    {sce_backlog_rate}%
                </div>
                """, unsafe_allow_html=True)

        with col3:
            compl_rate = results['performance_metrics']['completion_metrics']['completion_rate']
            st.markdown(
                f"""
                <div style="font-family: 'Tw Cen MT'; font-size: 18px; font-weight: bold; color: black;">
                    Completion Progress
                </div>
                <div style="font-family: 'Tw Cen MT'; font-size: 18px; font-weight: bold; color: green;">
                    {compl_rate}%
                </div>
                """, unsafe_allow_html=True)
            
        with col4:
            ytd_value = results['performance_metrics']['completion_metrics']['ytd_percentage']
            st.markdown(
                f"""
                <div style="font-family: 'Tw Cen MT'; font-size: 18px; font-weight: bold; color: black;">
                    YTD Progress
                </div>
                <div style="font-family: 'Tw Cen MT'; font-size: 18px; font-weight: bold; color: #1E90FF;">
                    {ytd_value}%
                </div>
                """, unsafe_allow_html=True)
                    
        tab1, tab2 = st.tabs(["Performance Analysis", "Backlog Analysis"])
        with tab1:
            st.plotly_chart(
                create_monthly_performance_chart(results['performance_metrics']['monthly_performance']),
                use_container_width=True)
            st.plotly_chart(
                create_monthly_performance_chart(results['sce_metrics']['monthly_performance'],
                                              title="SCE Monthly Performance Overview"),
                use_container_width=True)
            st.plotly_chart(
                create_completion_bar_chart(df, processor),  # Pass both df and processor
                use_container_width=True)

            st.markdown("""
                <style>
                div[data-baseweb="select"] {
                    width: 100px !important;
                }
                </style>
                """, unsafe_allow_html=True)

            selected_month = st.selectbox(
                "Monthly Progress",
                options=range(1, 13),
                format_func=lambda x: processor.month_map[x])

            monthly_performance = processor.get_monthly_item_class_performance(df, selected_month)

            def style_performance(val):
                target = val['Monthly Target']
                progress = val['Progress Accomplished']
                difference = target - progress
                if progress >= target:
                    return ['background-color: #90EE90']*len(val)
                elif 0 < difference <= 5:
                    return ['background-color: #FFB366']*len(val)
                else:
                    return ['background-color: #FF9999']*len(val)

            st.dataframe(
                monthly_performance.style.apply(style_performance, axis=1)
                .format("{:,.0f}")
                .set_properties(**{'text-align': 'center', 'color': 'black', 'font-weight': '500'})
                .set_table_styles([
                    {'selector': 'th', 'props': [('text-align', 'center'), ('color', 'black'), 
                                               ('font-weight', 'bold'), ('font-family', 'Tw Cen MT'), 
                                               ('font-size', '12px')]},
                    {'selector': 'td', 'props': [('text-align', 'center'), ('color', 'black'), 
                                               ('font-family', 'Tw Cen MT'), ('font-size', '12px')]}]))

        with tab2:
            st.subheader("Backlog Summary")
            df_backlog = pd.DataFrame.from_dict(results['item_class_analysis']['pivot_table'])
            df_backlog['Totals'] = df_backlog.sum(axis=1)
            grand_total = df_backlog.sum()
            df_backlog.loc['Grand Total'] = grand_total

            st.dataframe(
                df_backlog.style
                .format(lambda x: "" if pd.isna(x) or x == 0 else f"{int(x)}")
                .apply(lambda x: [apply_delay_colors(v, k) for v, k in zip(x, x.index)])
                .set_properties(**{
                    'text-align': 'center',
                    'color': 'black',
                    'font-family': 'Tw Cen MT',
                    'font-size': '9px'
                })
                .set_table_styles([
                    {'selector': 'th', 'props': [
                        ('text-align', 'center'),
                        ('color', 'black'),
                        ('font-family', 'Tw Cen MT'),
                        ('font-size', '9px'),
                        ('font-weight', 'bold')
                    ]},
                    {'selector': 'td', 'props': [
                        ('text-align', 'center'),
                        ('color', 'black'),
                        ('font-family', 'Tw Cen MT'),
                        ('font-size', '9px')
                    ]},
                    {'selector': 'td:last-child', 'props': [('font-weight', 'bold')]}
                ])
            )

            st.plotly_chart(create_backlog_item_class_chart(df_backlog), use_container_width=True)

            delay_options = list(processor.delay_colors.keys())
            selected_delay = st.selectbox("Filter by Delay", options=delay_options, key="delay_filter")

            if selected_delay:
                delay_details = processor.get_backlog_details_by_delay(df, selected_delay)
                st.dataframe(
                    delay_details.style
                    .set_properties(**{'text-align': 'center', 'color': 'black', 'font-weight': '500'})
                    .set_table_styles([
                        {'selector': 'th', 'props': [('text-align', 'center'), ('color', 'black'), 
                                                   ('font-family', 'Tw Cen MT'), ('font-size', '12px')]},
                        {'selector': 'td', 'props': [('text-align', 'left'), ('color', 'black'), 
                                                   ('font-family', 'Tw Cen MT'), ('font-size', '12px')]}]))

    else:
        st.info("Please upload the Insp Program file according to the site selection to begin the analysis.")

if __name__ == "__main__":
    main()