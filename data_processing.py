import pandas as pd
from datetime import datetime


class DataProcessor:
   def __init__(self):
       self.month_map = {
           1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR',
           5: 'MAY', 6: 'JUN', 7: 'JUL', 8: 'AUG',
           9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC'
       }
       self.delay_colors = {
           '< 6 Months': '#90EE90',  # Light green
           '6 Months < x <1 Yrs': '#FFB366', # Light orange
           '1 Yrs < x <2 Yrs': '#FFE5B4',  # Light yellow
           '2 Yrs < x <3 Yrs': '#FFC0CB',  # Light pink
           '> 3 Yrs': '#FF6B6B'  # Red
       }
       
   def load_inspection_data(self, uploaded_file, selected_site='GIR'):
       site_config = {
           'GIR': {'nrows': 583},
           'DAL': {'nrows': 761},
           'PAZ': {'nrows': 861}, 
           'CLV': {'nrows': 735}
       }

       # Add this before the main df = pd.read_excel()
       temp_df = pd.read_excel(
           uploaded_file,
           sheet_name='Data Base',
           skiprows=4,
           nrows=1
       )
       print("Available columns:", list(temp_df.columns))
       
       df = pd.read_excel(
           uploaded_file,
           sheet_name='Data Base',
           skiprows=4,
           nrows=site_config[selected_site]['nrows'],
           usecols=['Item Class', 'Backlog?', 'SECE STATUS', 'Delay', 'Year', 
                   'Job Done', 'CMonth Insp', 'PMonth Insp', 'Unit name', 'Scope']
       )
       
       df.columns = df.columns.str.strip()
       df = df.dropna(how='all')
       df['Backlog?'] = df['Backlog?'].fillna('No')
       df['SECE STATUS'] = df['SECE STATUS'].fillna('Non-SCE').astype(str).apply(
           lambda x: 'SCE' if x.upper() == 'SECE' else 'Non-SCE'
       )
       df['Year'] = df['Year'].fillna(datetime.now().year)
       df['Job Done'] = df['Job Done'].fillna('Not Compl')
       
       return df
   
   def get_backlog_details_by_delay(self, df, selected_delay):
       """Get details of backlog items for selected delay"""
       backlog_items = df[
           (df['Backlog?'] == 'Yes') & 
           (df['Delay'] == selected_delay)
       ][['Item Class', 'Unit name', 'Scope', 'SECE STATUS']]
       return backlog_items

   def analyze_data(self, df):
       backlog_analysis = self._analyze_backlog(df)
       performance_metrics = self._analyze_performance(df)
       sce_metrics = self._analyze_sce_performance(df)
       item_class_analysis = self._analyze_item_class_progress(df)
       
       return {
           'backlog_summary': backlog_analysis,
           'performance_metrics': performance_metrics,
           'sce_metrics': sce_metrics,
           'item_class_analysis': item_class_analysis
       }
   

   

    
    



   def _analyze_backlog(self, df):
       backlog_items = df[df['Backlog?'] == 'Yes']
       total_backlog = len(backlog_items)
       
       sece_backlog = len(backlog_items[backlog_items['SECE STATUS'] == 'SECE'])
       sece_percentage = (sece_backlog / total_backlog * 100) if total_backlog > 0 else 0
       
       return {
           'total_backlog': total_backlog,
           'sece_metrics': {
               'sece_backlog': sece_backlog,
               'sece_percentage': round(sece_percentage, 1)
           }
       }
   
   def _analyze_performance(self, df):
        df['PMonth_Num'] = pd.to_numeric(df['PMonth Insp'], errors='coerce')
        
        monthly_data = []
        today = datetime.now()
        current_month = today.month  # Get current month
        
        # Calculate backlog carryover
        previous_months_backlog = df[
            (df['PMonth_Num'] < current_month) & 
            (df['Backlog?'] == 'Yes')
        ]
        current_month_backlog = df[
            (df['PMonth_Num'] == current_month) & 
            (df['Backlog?'] == 'Yes')
        ]
        
        for month_num in range(1, 13):
            month_items = df[df['PMonth_Num'] == month_num]
            total_planned = len(month_items)
            completed_count = len(month_items[month_items['Job Done'] == 'Compl'])
            
            # Freeze backlog for terminated months (previous months)
            if month_num < current_month:
                backlog_count = len(previous_months_backlog[previous_months_backlog['PMonth_Num'] == month_num])
            # Carry backlog for the current month
            elif month_num == current_month:
                backlog_count = len(current_month_backlog) + len(previous_months_backlog)
            # Future months have no backlog
            else:
                backlog_count = 0
            
            progress_pct = (completed_count / total_planned * 100) if total_planned > 0 else 0
            
            monthly_data.append({
                'month': self.month_map[month_num],
                'total_planned': total_planned,
                'backlog_count': backlog_count,
                'completed_count': completed_count,
                'progress_percentage': round(progress_pct, 1)
            })
            
        return {
            'monthly_performance': monthly_data,
            'completion_metrics': self._calculate_completion_metrics(df)
        }

   def _analyze_sce_performance(self, df):
        sce_df = df[df['SECE STATUS'] == 'SECE']
        return self._analyze_performance(sce_df)

   def _analyze_item_class_progress(self, df):
       backlog_items = df[df['Backlog?'] == 'Yes']
       
       # Define the correct order for Delay categories
       delay_order = list(self.delay_colors.keys())
       
       # Define SECE STATUS order
       sece_order = ['SECE', 'Non-SCE']
       
       # Create pivot table
       pivot_table = pd.pivot_table(
           backlog_items,
           values='Backlog?',
           index='Item Class',
           columns=['Delay', 'SECE STATUS'],
           aggfunc='count',
           fill_value=0
       )
       
       # Reorder columns according to specified order
       ordered_columns = pd.MultiIndex.from_product(
           [delay_order, sece_order],
           names=['Delay', 'SECE STATUS']
       )
       pivot_table = pivot_table.reindex(columns=ordered_columns)
       
       total_backlog = len(backlog_items)
       sece_backlog = len(backlog_items[backlog_items['SECE STATUS'] == 'SECE'])
       equipment_classes = len(pivot_table.index)
       
       return {
           'pivot_table': pivot_table.to_dict(),
           'summary_stats': {
               'total_backlog': total_backlog,
               'sece_backlog': sece_backlog,
               'equipment_classes': equipment_classes
           }
       }

   def _calculate_completion_metrics(self, df):
        df['CMonth_Num'] = pd.to_numeric(df['CMonth Insp'], errors='coerce')
        df['PMonth_Num'] = pd.to_numeric(df['PMonth Insp'], errors='coerce')
        
        total_jobs = len(df)
        completed_jobs = len(df[df['Job Done'] == 'Compl'])
        completion_rate = (completed_jobs / total_jobs * 100) if total_jobs > 0 else 0
        
        on_time_jobs = len(df[
            (df['Job Done'] == 'Compl') & 
            (df['CMonth_Num'] <= df['PMonth_Num'])
        ])
        on_time_rate = (on_time_jobs / completed_jobs * 100) if completed_jobs > 0 else 0
        
        # Calculate YTD percentage based on weeks from Jan 1, 2025
        today = datetime.now()
        jan_1_2025 = datetime(2025, 1, 1)  # Start date for YTD calculation
        delta = today - jan_1_2025
        week_number = today.isocalendar()[1] #delta.days // 7  # Calculate weeks since Jan 1, 2025
        ytd_percentage = (week_number / 52) * 100  # YTD percentage based on 52 weeks
        
        # Round to nearest whole integer and ensure it's even
        ytd_percentage_rounded = round(ytd_percentage)
        if ytd_percentage_rounded % 2 != 0:  # If odd, round to nearest even
            ytd_percentage_rounded = round(ytd_percentage / 2) * 2
        
        return {
            'total_jobs': total_jobs,
            'completed_jobs': completed_jobs,
            'completion_rate': round(completion_rate, 1),
            'on_time_rate': round(on_time_rate, 1),
            'ytd_percentage': int(ytd_percentage_rounded)  # Add YTD percentage (whole even integer)
        }
    
   def get_monthly_item_class_performance(self, df, selected_month):
    """Analyze item class performance for selected month"""
    df['PMonth_Num'] = pd.to_numeric(df['PMonth Insp'], errors='coerce')
    
    # Filter for selected month
    month_data = df[df['PMonth_Num'] == selected_month]
    
    # Group by Item Class
    performance_data = pd.DataFrame()
    performance_data['2025 SOW'] = df.groupby('Item Class').size()
    performance_data['Monthly Target'] = month_data.groupby('Item Class').size()
    performance_data['Progress Accomplished'] = month_data[month_data['Job Done'] == 'Compl'].groupby('Item Class').size()
    
    # Fill NaN with 0
    performance_data = performance_data.fillna(0).astype(int)
    

    return performance_data





