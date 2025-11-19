import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
from datetime import datetime

# Define fixed colors for rejection reasons
REJECTION_REASON_COLORS = {
    'Assembly': '#ff9999',          # Light red
    'Optical quality': '#66b3ff',  # Light blue
    'Injection failure': '#99ff99', # Light green
    'Human error': '#f7d560',       # Light red
    'Not sealed': '#fa5ff7',        # Light purple
    'Faild accommodation': '#ff796c', # Light yellow
    'Other': '#ffb366'              # Orange (last as you wanted)
}


# Page configuration
st.set_page_config(page_title="AIOL Production Rejection Analysis", page_icon="🔍", layout="wide")
st.title("🔍 AIOL Production Rejection Reasons Statistics")
st.markdown('<style>div.block-container{padding-top:1rem;}</style>', unsafe_allow_html=True)

# File upload - allow multiple files
uploaded_files = st.file_uploader(
    ":file_folder: Upload Excel files (AIOL Production Data)",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)


def extract_batch_info(filename):
    """Extract batch information from filename - looking for SN pattern like 201-230"""
    try:
        # Look for pattern like "201-230" or "SN201-230" in filename
        match = re.search(r'(?:SN\s*)?(\d{3,4})-(\d{3,4})', filename)

        if match:
            start_sn = int(match.group(1))
            end_sn = int(match.group(2))
            return f"SN{start_sn}-{end_sn}", start_sn, end_sn

        # Fallback: try to find it in any format
        match_alt = re.search(r'(\d{3,4})\s*-\s*(\d{3,4})', filename)
        if match_alt:
            start_sn = int(match_alt.group(1))
            end_sn = int(match_alt.group(2))
            return f"SN{start_sn}-{end_sn}", start_sn, end_sn

        return "Unknown Batch", 9999, 9999  # High numbers for sorting unknown batches last
    except Exception as e:
        st.warning(f"Could not extract batch info from filename: {str(e)}")
        return "Unknown Batch", 9999, 9999


def parse_status_data(df, batch_name):
    """Parse the status data from column A starting from row 6 and serial numbers from column C"""
    try:
        # Get data from column A starting from row 6 (index 5)
        status_data = df.iloc[4:, 0].dropna()  # Column A, starting from row 6

        # Get serial numbers from column C starting from row 6 (index 5, column 2)
        serial_data = df.iloc[4:, 2] if df.shape[1] > 2 else pd.Series()  # Column C, starting from row 6

        parsed_data = []
        other_statuses = set()  # Track unexpected statuses

        for idx, value in status_data.items():
            value_str = str(value).strip()

            # Get corresponding serial number
            serial_number = ''
            if idx < len(serial_data) + 4:  # Adjust for 0-based indexing
                serial_value = serial_data.iloc[idx - 4] if not serial_data.empty and (idx - 4) < len(
                    serial_data) else None
                if pd.notna(serial_value):
                    serial_number = str(serial_value).strip()

            if value_str and value_str != 'nan':
                # Handle different status cases
                value_lower = value_str.lower()

                if value_lower == 'approved':
                    parsed_data.append({
                        'Batch': batch_name,
                        'AIOL_Serial_Number': serial_number,
                        'Status': 'Approved',
                        'Rejection_Reason': '',
                        'Subfield': '',
                        'Raw_Data': value_str
                    })

                elif value_lower == 'in process':
                    # Don't count in process items, but track them
                    continue

                elif value_str.lower().startswith('rejected'):
                    # Parse rejected items
                    parts = re.split(r'\s*-\s*', value_str)

                    if len(parts) >= 2:
                        status = 'Rejected'

                        if len(parts) == 2:
                            rejection_reason = parts[1].strip()
                            subfield = 'No subfield'
                        else:
                            rejection_reason = parts[1].strip()
                            subfield = parts[2].strip() if len(parts) >= 3 else 'No subfield'

                        # CHANGE: Treat assembly rejections as 'In Process'
                        if rejection_reason.lower() == 'assembly':
                            # Skip assembly rejections - treat as in process
                            continue

                        parsed_data.append({
                            'Batch': batch_name,
                            'AIOL_Serial_Number': serial_number,
                            'Status': status,
                            'Rejection_Reason': rejection_reason,
                            'Subfield': subfield,
                            'Raw_Data': value_str
                        })
                    else:
                        # Handle malformed rejected entries (but skip if assembly)
                        if 'assembly' not in value_str.lower():
                            parsed_data.append({
                                'Batch': batch_name,
                                'AIOL_Serial_Number': serial_number,
                                'Status': 'Rejected',
                                'Rejection_Reason': 'Unknown',
                                'Subfield': 'No subfield',
                                'Raw_Data': value_str
                            })
                else:
                    # Track other unexpected statuses
                    other_statuses.add(value_str)

        return parsed_data, other_statuses

    except Exception as e:
        st.error(f"Error parsing status data for {batch_name}: {str(e)}")
        return [], set()

def standardize_rejection_reasons(df):
    """Standardize rejection reasons - map injection failure to Other"""
    if 'Rejection_Reason' in df.columns:
        df['Rejection_Reason'] = df['Rejection_Reason'].replace({
            'Injection failure': 'Other',
            'injection failure': 'Other'
        })
    return df

def load_and_process_files(files):
    """Load and process all uploaded Excel files"""
    all_data = []
    all_other_statuses = set()
    batch_info = {}

    for file in files:
        try:
            # Extract batch information from FILENAME
            batch_name, start_sn, end_sn = extract_batch_info(file.name)
            batch_info[file.name] = {
                'batch_name': batch_name,
                'start_sn': start_sn,
                'end_sn': end_sn
            }

            # Read the Excel file
            df = pd.read_excel(file, header=None, engine='openpyxl')

            # Parse the status data
            parsed_data, other_statuses = parse_status_data(df, batch_name)

            # Convert to dataframe, standardize, then extend
            if parsed_data:
                temp_df = pd.DataFrame(parsed_data)
                temp_df = standardize_rejection_reasons(temp_df)
                all_data.extend(temp_df.to_dict('records'))
            else:
                all_data.extend(parsed_data)

            all_other_statuses.update(other_statuses)

            st.success(f"✅ Processed {file.name} - Batch: {batch_name} - {len(parsed_data)} records")

        except Exception as e:
            st.error(f"❌ Error processing {file.name}: {str(e)}")

    return all_data, all_other_statuses, batch_info


# Main processing
if uploaded_files:
    st.write(f"📁 Processing {len(uploaded_files)} file(s)...")

    # Load and process all files
    all_data, other_statuses, batch_info = load_and_process_files(uploaded_files)

    if other_statuses:
        st.warning(f"⚠️ Found unexpected statuses (not counted): {', '.join(other_statuses)}")

    if all_data:
        df_combined = pd.DataFrame(all_data)
        rejected_data = df_combined[df_combined['Status'] == 'Rejected']

        if not rejected_data.empty:
            # Define expected rejection reasons based on your new format
            expected_reasons = [
                'Assembly', 'Optical quality',
                'Human error', 'Not sealed', 'Faild accommodation', 'Other'
            ]

            # Find rejection reasons that don't match expected ones (case-insensitive)
            unexpected_reasons = set()
            expected_reasons_lower = [r.lower() for r in expected_reasons]

            for reason in rejected_data['Rejection_Reason'].unique():
                if pd.notna(reason):
                    reason_lower = str(reason).lower().strip()
                    if reason_lower not in expected_reasons_lower:
                        unexpected_reasons.add(str(reason))

            if unexpected_reasons:
                st.warning(f"⚠️ Found unexpected rejection reasons (check for typos): {', '.join(unexpected_reasons)}")

    if all_data:
        # Convert to DataFrame
        df_combined = pd.DataFrame(all_data)

        st.success(f"📊 Total processed records: {len(df_combined)}")

        # Display batch information
        st.write("### Batch Information")
        batch_display = []
        for file, info in batch_info.items():
            batch_display.append({
                'File': file,
                'Batch': info['batch_name'],
                'Start_SN': info['start_sn'],
                'End_SN': info['end_sn']
            })
        batch_df = pd.DataFrame(batch_display)
        # Sort by Start_SN
        batch_df = batch_df.sort_values('Start_SN')
        # Display without the SN columns (they're just for sorting)
        st.dataframe(batch_df[['File', 'Batch']], use_container_width=True)
        # Create sorted list of batches for filtering
        sorted_batches = batch_df['Batch'].tolist()

        #batch_df = pd.DataFrame(list(batch_info.items()), columns=['File', 'Batch'])
        #st.dataframe(batch_df, use_container_width=True)

        # Sidebar filters
        st.sidebar.header("Filter Options")

        # Batch selection
        available_batches = df_combined['Batch'].unique()
        selected_batches = st.sidebar.multiselect(
            "Select Batches to Analyze",
            sorted_batches,
            default=sorted_batches
        )

        # Visualization type selection
        combine_batches = st.sidebar.checkbox(
            "Combine all selected batches",
            value=True,
            help="When checked: shows statistics for all selected batches together. When unchecked: shows each batch separately."
        )

        # Filter data based on selection
        if selected_batches:
            filtered_df = df_combined[df_combined['Batch'].isin(selected_batches)]
        else:
            filtered_df = df_combined

        if not filtered_df.empty:
            # Calculate statistics
            total_items = len(filtered_df)
            approved_count = len(filtered_df[filtered_df['Status'] == 'Approved'])
            rejected_count = len(filtered_df[filtered_df['Status'] == 'Rejected'])

            # Display summary statistics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Items", total_items)
            with col2:
                st.metric("Approved", approved_count)
            with col3:
                st.metric("Rejected", rejected_count)
            with col4:
                rejection_rate = (rejected_count / total_items * 100) if total_items > 0 else 0
                st.metric("Rejection Rate", f"{rejection_rate:.1f}%")

            # Create visualizations based on selected view type
            if combine_batches:
                st.subheader("📊 Combined Analysis (All Selected Batches)")

                # Combined status bar chart
                col1, col2 = st.columns(2)

                with col1:
                    # Single stacked bar chart for overall status
                    status_counts = filtered_df['Status'].value_counts()
                    total_items = len(filtered_df)

                    # Calculate percentages
                    status_percentages = (status_counts / total_items * 100).round(1)

                    fig_status = go.Figure()

                    # Add approved and rejected bars with specific colors
                    if 'Approved' in status_counts.index:
                        approved_count = status_counts['Approved']
                        approved_pct = status_percentages['Approved']
                        fig_status.add_trace(go.Bar(
                            name='Approved',
                            x=['Status Distribution'],
                            y=[approved_pct],
                            marker_color='#28a745',  # Green
                            text=f'{approved_pct:.1f}% ({approved_count})',
                            textposition='inside',
                            textfont=dict(color='white', size=12),
                            hovertemplate='<b>Approved</b><br>Percentage: %{y:.1f}%<br>Count: ' + str(
                                approved_count) + '<extra></extra>'
                        ))

                    if 'Rejected' in status_counts.index:
                        rejected_count = status_counts['Rejected']
                        rejected_pct = status_percentages['Rejected']
                        fig_status.add_trace(go.Bar(
                            name='Rejected',
                            x=['Status Distribution'],
                            y=[rejected_pct],
                            marker_color='#dc3545',  # Red
                            text=f'{rejected_pct:.1f}% ({rejected_count})',
                            textposition='inside',
                            textfont=dict(color='white', size=12),
                            hovertemplate='<b>Rejected</b><br>Percentage: %{y:.1f}%<br>Count: ' + str(
                                rejected_count) + '<extra></extra>'
                        ))

                    fig_status.update_layout(
                        title="Approval/Rejection Status Distribution (%)",
                        barmode='stack',
                        xaxis_title='',
                        yaxis_title='Percentage (%)',
                        showlegend=True,
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="center",
                            x=0.5
                        )
                    )

                    st.plotly_chart(fig_status, use_container_width=True)

                with col2:
                    # Single stacked bar chart for rejection reasons (percentages of total items)
                    rejected_df = filtered_df[filtered_df['Status'] == 'Rejected']
                    total_items = len(filtered_df)  # FIXED: Back to total items for proper percentage
                    if not rejected_df.empty:
                        rejection_counts = rejected_df['Rejection_Reason'].value_counts()
                        rejection_percentages = (rejection_counts / total_items * 100).round(1)

                        # Create a single stacked bar chart
                        fig_bar = go.Figure()

                        # Define colors for different rejection reasons
                        #colors = px.colors.qualitative.Set3[:len(rejection_percentages)]

                        # Add each rejection reason as a segment in the stacked bar
                        for i, (reason, percentage) in enumerate(rejection_percentages.items()):
                            count = rejection_counts[reason]
                            fig_bar.add_trace(go.Bar(
                                name=reason,
                                x=['Rejection Reasons'],
                                y=[percentage],
                                #marker_color=colors[i % len(colors)],
                                marker_color=REJECTION_REASON_COLORS.get(reason, '#cccccc'),
                                text=f'{percentage:.1f}% ({count})',
                                textposition='inside',
                                textfont=dict(color='black', size=12),
                                hovertemplate=f'<b>{reason}</b><br>Percentage: {percentage:.1f}%<br>Count: {count}<extra></extra>'
                            ))
                            # Reorder legend to put "Other" last
                            fig_bar.update_layout(
                                legend=dict(
                                    traceorder="normal"
                                )
                            )

                            # Get current traces and reorder them
                            traces = list(fig_bar.data)
                            other_traces = [t for t in traces if t.name.lower() == 'other']
                            non_other_traces = [t for t in traces if t.name.lower() != 'other']

                            # Clear and re-add traces in new order
                            fig_bar.data = []
                            for trace in non_other_traces + other_traces:
                                fig_bar.add_trace(trace)

                        fig_bar.update_layout(
                            title="Rejection Reasons Distribution (%)",
                            barmode='stack',
                            xaxis_title='',
                            yaxis_title='Percentage (%)',
                            showlegend=True,
                            legend=dict(
                                orientation="v",
                                yanchor="top",
                                y=1,
                                xanchor="left",
                                x=1.02
                            )
                        )

                        st.plotly_chart(fig_bar, use_container_width=True)
                    else:
                        st.info("No rejected items in selected data")


                # NEW: Combined pie chart showing Accepted + all rejection reasons
                st.write("#### Overall Status with Rejection Breakdown")

                # Calculate counts for pie chart
                pie_data = {'Accepted': approved_count}

                # Add each rejection reason
                rejected_df = filtered_df[filtered_df['Status'] == 'Rejected']
                if not rejected_df.empty:
                    rejection_counts = rejected_df['Rejection_Reason'].value_counts()
                    for reason, count in rejection_counts.items():
                        pie_data[reason] = count

                # Create colors list - green for accepted, then rejection reason colors
                pie_colors = []
                pie_labels = []
                pie_values = []

                for label, value in pie_data.items():
                    pie_labels.append(label)
                    pie_values.append(value)
                    if label == 'Accepted':
                        pie_colors.append('#28a745')  # Green
                    else:
                        pie_colors.append(REJECTION_REASON_COLORS.get(label, '#cccccc'))

                fig_pie_combined = go.Figure(data=[go.Pie(
                    labels=pie_labels,
                    values=pie_values,
                    marker=dict(colors=pie_colors),
                    textinfo='label+percent',
                    texttemplate='%{label}<br>%{percent}%',
                    hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent:.1f}%<extra></extra>'
                )])

                fig_pie_combined.update_layout(
                    title="Accepted vs Rejected Breakdown",
                    showlegend=True
                )

                st.plotly_chart(fig_pie_combined, use_container_width=True)
                # NEW: Pie chart for rejection reasons only
                st.write("#### Rejection Reasons Distribution")

                rejected_df = filtered_df[filtered_df['Status'] == 'Rejected']
                if not rejected_df.empty:
                    rejection_counts = rejected_df['Rejection_Reason'].value_counts()

                    # Create colors for rejection reasons
                    rejection_labels = []
                    rejection_values = []
                    rejection_colors = []

                    for reason, count in rejection_counts.items():
                        rejection_labels.append(reason)
                        rejection_values.append(count)
                        rejection_colors.append(REJECTION_REASON_COLORS.get(reason, '#cccccc'))

                    fig_rejection_pie = go.Figure(data=[go.Pie(
                        labels=rejection_labels,
                        values=rejection_values,
                        marker=dict(colors=rejection_colors),
                        textinfo='label+percent',
                        texttemplate='%{label}<br>%{percent}%',
                        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent:.1f}%<extra></extra>'
                    )])

                    fig_rejection_pie.update_layout(
                        title="Rejection Reasons Only",
                        showlegend=True
                    )

                    st.plotly_chart(fig_rejection_pie, use_container_width=True)
                    # View data option

                    with st.expander("📋 View Rejection Reasons Data"):
                        # Include both accepted and rejected
                        status_summary = filtered_df.groupby('Status').size().reset_index(name='Count')

                        # Add rejection reasons breakdown
                        rejected_df = filtered_df[filtered_df['Status'] == 'Rejected']
                        if not rejected_df.empty:
                            rejection_summary = rejected_df.groupby('Rejection_Reason').size().reset_index(name='Count')
                            rejection_summary['Status'] = rejection_summary['Rejection_Reason']

                            # Combine accepted with rejection reasons
                            accepted_count = len(filtered_df[filtered_df['Status'] == 'Accepted'])
                            accepted_row = pd.DataFrame({'Status': ['Accepted'], 'Count': [accepted_count]})

                            combined_summary = pd.concat([accepted_row, rejection_summary[['Status', 'Count']]],
                                                         ignore_index=True)
                        else:
                            combined_summary = status_summary

                        # Calculate percentage of total items
                        total_items = len(filtered_df)
                        combined_summary['Percentage'] = (combined_summary['Count'] / total_items * 100).round(2)

                        st.dataframe(combined_summary, use_container_width=True)

                        csv_rejection = combined_summary.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            "📥 Download Rejection Reasons CSV",
                            csv_rejection,
                            "rejection_reasons.csv",
                            "text/csv",
                            key='download-rejection-reasons-combined'
                        )

                        st.dataframe(rejection_summary, use_container_width=True)

                        csv_rejection = rejection_summary.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            "📥 Download Rejection Reasons CSV",
                            csv_rejection,
                            "rejection_reasons.csv",
                            "text/csv",
                            key='download-rejection-reasons'
                        )


                # Optical quality subfields analysis
                st.write("### Optical Quality Rejection Analysis")
                optical_rejected = filtered_df[(filtered_df['Status'] == 'Rejected') &
                                               (filtered_df['Rejection_Reason'].str.lower() == 'optical quality')]

                if not optical_rejected.empty:
                    # Single stacked bar chart for optical quality subfields (percentages of total items)
                    total_items = len(filtered_df)  # FIXED: Back to total items
                    subfield_counts = optical_rejected['Subfield'].value_counts()
                    subfield_percentages = (subfield_counts / total_items * 100).round(1)

                    # Create a single stacked bar chart for optical quality subfields
                    fig_optical = go.Figure()

                    # Define colors for different subfields (different palette from assembly)
                    colors = px.colors.qualitative.Set2[:len(subfield_percentages)]

                    # Add each subfield as a segment in the stacked bar
                    for i, (subfield, percentage) in enumerate(subfield_percentages.items()):
                        count = subfield_counts[subfield]
                        fig_optical.add_trace(go.Bar(
                            name=subfield,
                            x=['Optical Quality Rejection Subfields'],
                            y=[percentage],
                            marker_color=colors[i % len(colors)],
                            text=f' {percentage:.1f}% ({count})',
                            textposition='inside',
                            textfont=dict(color='black', size=12),
                            hovertemplate=f'<b>{subfield}</b><br>Percentage: {percentage:.1f}%<br>Count: {count}<extra></extra>'
                        ))

                    fig_optical.update_layout(
                        title="Optical Quality Rejection Distribution (%)",
                        barmode='stack',
                        xaxis_title='',
                        yaxis_title='Percentage (%)',
                        showlegend=True,
                        legend=dict(
                            orientation="v",
                            yanchor="top",
                            y=1,
                            xanchor="left",
                            x=1.02
                        )
                    )

                    st.plotly_chart(fig_optical, use_container_width=True)
                    # Combined pie chart for Optical Quality - Accepted + Subfields
                    st.write("#### Optical Quality - Accepted vs Subfield Breakdown")

                    # Calculate counts: all accepted + optical quality subfields
                    total_items = len(filtered_df)
                    accepted_count = len(filtered_df[filtered_df['Status'] == 'Accepted'])

                    # Get optical quality subfield counts
                    optical_pie_data = {'Accepted': accepted_count}
                    if not optical_rejected.empty:
                        subfield_counts = optical_rejected['Subfield'].value_counts()
                        for subfield, count in subfield_counts.items():
                            optical_pie_data[subfield] = count

                    # Create colors - green for accepted, then Set2 palette for subfields
                    optical_pie_colors = []
                    optical_pie_labels = []
                    optical_pie_values = []

                    # Define colors for optical subfields (using Set2 palette)
                    subfield_color_map = {}
                    subfield_palette = px.colors.qualitative.Set2
                    subfield_idx = 0

                    for label, value in optical_pie_data.items():
                        optical_pie_labels.append(label)
                        optical_pie_values.append(value)
                        if label == 'Accepted':
                            optical_pie_colors.append('#28a745')  # Green
                        else:
                            if label not in subfield_color_map:
                                subfield_color_map[label] = subfield_palette[subfield_idx % len(subfield_palette)]
                                subfield_idx += 1
                            optical_pie_colors.append(subfield_color_map[label])

                    fig_optical_pie_combined = go.Figure(data=[go.Pie(
                        labels=optical_pie_labels,
                        values=optical_pie_values,
                        marker=dict(colors=optical_pie_colors),
                        textinfo='label+percent',
                        texttemplate='%{label}<br>%{percent}',
                        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
                    )])

                    fig_optical_pie_combined.update_layout(
                        title="Accepted vs Optical Quality Subfield Breakdown",
                        showlegend=True
                    )

                    st.plotly_chart(fig_optical_pie_combined, use_container_width=True)

                    # Show optical quality rejection summary
                    st.info(
                        f"📊 Optical quality rejections: {len(optical_rejected)} out of {len(filtered_df[filtered_df['Status'] == 'Rejected'])} total rejections")
                    # View data option
                    with st.expander("📋 View Optical Quality Rejection Data"):
                        optical_summary = optical_rejected.groupby('Subfield').size().reset_index(name='Count')
                        optical_summary['Percentage'] = (optical_summary['Count'] / len(filtered_df) * 100).round(2)
                        st.dataframe(optical_summary, use_container_width=True)

                        # Also show detailed records
                        st.write("**Detailed Records:**")
                        st.dataframe(optical_rejected[['Batch', 'AIOL_Serial_Number', 'Subfield']],
                                     use_container_width=True)

                        csv_optical = optical_rejected.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            "📥 Download Optical Quality Details CSV",
                            csv_optical,
                            "optical_quality_details.csv",
                            "text/csv",
                            key='download-optical-quality-combined'
                        )
                else:
                    st.info("No optical quality rejections found in selected data")

            else:
                st.subheader("📊 Batch-by-Batch Analysis")

                # Single stacked bar chart for status distribution by batch
                st.write("#### Approval/Rejection Status Distribution by Batch (%)")

                # Calculate status percentages for each batch
                batch_status_data = []

                for batch in selected_batches:
                    batch_data = filtered_df[filtered_df['Batch'] == batch]

                    if not batch_data.empty:
                        batch_total = len(batch_data)
                        status_counts = batch_data['Status'].value_counts()

                        for status in ['Approved', 'Rejected']:
                            if status in status_counts.index:
                                count = status_counts[status]
                                percentage = (count / batch_total * 100)
                                batch_status_data.append({
                                    'Batch': batch,
                                    'Status': status,
                                    'Percentage': percentage,
                                    'Count': count
                                })
                            else:
                                batch_status_data.append({
                                    'Batch': batch,
                                    'Status': status,
                                    'Percentage': 0,
                                    'Count': 0
                                })

                if batch_status_data:
                    status_df = pd.DataFrame(batch_status_data)

                    # Create stacked bar chart with one bar per batch
                    fig_status_batch = go.Figure()

                    # Add approved bars (green)
                    approved_data = status_df[status_df['Status'] == 'Approved']
                    if not approved_data.empty:
                        fig_status_batch.add_trace(go.Bar(
                            name='Approved',
                            x=approved_data['Batch'],
                            y=approved_data['Percentage'],
                            marker_color='#28a745',  # Green
                            text=[f'{row.Percentage:.1f}% ({row.Count})' for _, row in approved_data.iterrows()],
                            textposition='inside',
                            textfont=dict(color='white', size=12),
                            hovertemplate='<b>Approved</b><br>Percentage: %{y:.1f}%<br>Count: %{customdata}<extra></extra>',
                            customdata=approved_data['Count']
                        ))

                    # Add rejected bars (red)
                    rejected_data = status_df[status_df['Status'] == 'Rejected']
                    if not rejected_data.empty:
                        fig_status_batch.add_trace(go.Bar(
                            name='Rejected',
                            x=rejected_data['Batch'],
                            y=rejected_data['Percentage'],
                            marker_color='#dc3545',  # Red
                            text=[f'{row.Percentage:.1f}% ({row.Count})' for _, row in rejected_data.iterrows()],
                            textposition='inside',
                            textfont=dict(color='white', size=12),
                            hovertemplate='<b>Rejected</b><br>Percentage: %{y:.1f}%<br>Count: %{customdata}<extra></extra>',
                            customdata=rejected_data['Count']
                        ))

                    fig_status_batch.update_layout(
                        title='Approval/Rejection Status Distribution by Batch (%)',
                        barmode='stack',
                        xaxis_title='Batch',
                        yaxis_title='Percentage (%)',
                        legend_title='Status',
                        showlegend=True
                    )

                    st.plotly_chart(fig_status_batch, use_container_width=True)
                    # NEW CODE FOR BATCH VIEW - ADD HERE:
                    with st.expander("📋 View Batch Status Data"):
                        batch_status_summary = filtered_df.groupby(['Batch', 'Status']).size().reset_index(name='Count')
                        batch_totals = filtered_df.groupby('Batch').size().reset_index(name='Total')
                        batch_status_summary = batch_status_summary.merge(batch_totals, on='Batch')
                        batch_status_summary['Percentage'] = (
                                    batch_status_summary['Count'] / batch_status_summary['Total'] * 100).round(2)

                        st.dataframe(batch_status_summary, use_container_width=True)

                        csv_batch_status = batch_status_summary.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            "📥 Download Batch Status CSV",
                            csv_batch_status,
                            "batch_status.csv",
                            "text/csv",
                            key='download-batch-status'
                        )
                else:
                    st.info("No status data available for selected batches")

                # Rejection reasons bar chart - single stacked bar for each batch
                st.write("#### Rejection Reasons Distribution by Batch (%)")

                rejected_by_batch = filtered_df[filtered_df['Status'] == 'Rejected']

                if not rejected_by_batch.empty:
                    # Calculate rejection reason percentages for each batch
                    batch_rejection_data = []

                    for batch in selected_batches:
                        batch_rejected = rejected_by_batch[rejected_by_batch['Batch'] == batch]
                        batch_total = len(filtered_df[filtered_df['Batch'] == batch])  # FIXED: total items in batch

                        if not batch_rejected.empty:
                            reason_counts = batch_rejected['Rejection_Reason'].value_counts()

                            for reason, count in reason_counts.items():
                                percentage = (count / batch_total * 100)  # FIXED: percentage of total items
                                batch_rejection_data.append({
                                    'Batch': batch,
                                    'Rejection_Reason': reason,
                                    'Percentage': percentage,
                                    'Count': count
                                })

                    if batch_rejection_data:
                        rejection_df = pd.DataFrame(batch_rejection_data)

                        # Get all unique rejection reasons for consistent coloring
                        all_reasons = rejection_df['Rejection_Reason'].unique()
                        colors = px.colors.qualitative.Set3[:len(all_reasons)]
                        #color_map = {reason: colors[i] for i, reason in enumerate(all_reasons)}
                        color_map = REJECTION_REASON_COLORS  #  fixed color mapping

                        # Create stacked bar chart with one bar per batch
                        fig_rejection = go.Figure()

                        for reason in all_reasons:
                            reason_data = rejection_df[rejection_df['Rejection_Reason'] == reason]

                            # Create lists for batches and percentages
                            batch_list = []
                            percentage_list = []
                            text_list = []
                            count_list = []

                            for batch in selected_batches:
                                batch_reason_data = reason_data[reason_data['Batch'] == batch]
                                if not batch_reason_data.empty:
                                    percentage = batch_reason_data['Percentage'].iloc[0]
                                    count = int(batch_reason_data['Count'].iloc[0])
                                    batch_list.append(batch)
                                    percentage_list.append(percentage)
                                    text_list.append(f'{percentage:.1f}%')
                                    count_list.append(count)
                                else:
                                    batch_list.append(batch)
                                    percentage_list.append(0)
                                    text_list.append('')
                                    count_list.append(0)

                            fig_rejection.add_trace(go.Bar(
                                name=reason,
                                x=batch_list,
                                y=percentage_list,
                                #marker_color=color_map[reason],
                                marker_color=REJECTION_REASON_COLORS.get(reason, '#cccccc'),
                                text=text_list,
                                textposition='inside',
                                textfont=dict(color='black', size=12),
                                customdata=count_list,
                                hovertemplate=f'<b>{reason}</b><br>Percentage: %{{y:.1f}}%<br>Count: %{{customdata}}<extra></extra>'
                            ))

                        fig_rejection.update_layout(
                            title='Rejection Reasons by Batch (% of Total Items)',
                            barmode='stack',
                            xaxis_title='Batch',
                            yaxis_title='Percentage (%)',
                            legend_title='Rejection Reason',
                            showlegend=True
                        )
                        # Reorder legend to put "Other" last
                        fig_rejection.update_layout(
                            legend=dict(
                                traceorder="normal"  # or use a custom ordering function
                            )
                        )

                        # Get current traces and reorder them
                        traces = list(fig_rejection.data)
                        other_traces = [t for t in traces if t.name.lower() == 'other']
                        non_other_traces = [t for t in traces if t.name.lower() != 'other']

                        # Clear and re-add traces in new order
                        fig_rejection.data = []
                        for trace in non_other_traces + other_traces:
                            fig_rejection.add_trace(trace)

                        # Reorder legend to put "Other" last
                        fig_rejection.update_layout(
                            legend=dict(
                                traceorder="normal"  # or use a custom ordering function
                            )
                        )

                        # Get current traces and reorder them
                        traces = list(fig_rejection.data)
                        other_traces = [t for t in traces if t.name.lower() == 'other']
                        non_other_traces = [t for t in traces if t.name.lower() != 'other']

                        # Clear and re-add traces in new order
                        fig_rejection.data = []
                        for trace in non_other_traces + other_traces:
                            fig_rejection.add_trace(trace)

                        st.plotly_chart(fig_rejection, use_container_width=True)
                        # NEW: View data option with batch source
                        with st.expander("📋 View Batch Rejection Reasons Data"):
                            # Show rejection reasons by batch
                            rejected_by_batch = filtered_df[filtered_df['Status'] == 'Rejected']
                            batch_rejection_summary = rejected_by_batch.groupby(
                                ['Batch', 'Rejection_Reason']).size().reset_index(name='Count')

                            # Add batch totals for percentage calculation
                            batch_totals = filtered_df.groupby('Batch').size().reset_index(name='Total')
                            batch_rejection_summary = batch_rejection_summary.merge(batch_totals, on='Batch')
                            batch_rejection_summary['Percentage of Total'] = (
                                        batch_rejection_summary['Count'] / batch_rejection_summary[
                                    'Total'] * 100).round(2)

                            # Also calculate percentage of rejected items per batch
                            batch_rejected_totals = rejected_by_batch.groupby('Batch').size().reset_index(
                                name='Rejected_Total')
                            batch_rejection_summary = batch_rejection_summary.merge(batch_rejected_totals, on='Batch')
                            batch_rejection_summary['Percentage of Rejected'] = (
                                        batch_rejection_summary['Count'] / batch_rejection_summary[
                                    'Rejected_Total'] * 100).round(2)

                            # Display with relevant columns
                            display_cols = ['Batch', 'Rejection_Reason', 'Count', 'Percentage of Total',
                                            'Percentage of Rejected']
                            st.dataframe(batch_rejection_summary[display_cols], use_container_width=True)

                            csv_batch_rejection = batch_rejection_summary.to_csv(index=False).encode('utf-8')
                            st.download_button(
                                "📥 Download Batch Rejection Reasons CSV",
                                csv_batch_rejection,
                                "batch_rejection_reasons.csv",
                                "text/csv",
                                key='download-batch-rejection-reasons'
                            )
                    else:
                        st.info("No rejection data available for selected batches")
                else:
                    st.info("No rejected items in selected data")



                # Optical quality subfields analysis by batch
                st.write("#### Optical Quality Rejection Subfields by Batch (%)")

                optical_rejected_by_batch = filtered_df[(filtered_df['Status'] == 'Rejected') &
                                                        (filtered_df[
                                                             'Rejection_Reason'].str.lower() == 'optical quality')]

                if not optical_rejected_by_batch.empty:
                    # Calculate optical quality subfield percentages for each batch
                    batch_optical_data = []

                    for batch in selected_batches:
                        batch_optical = optical_rejected_by_batch[optical_rejected_by_batch['Batch'] == batch]
                        batch_total = len(filtered_df[filtered_df['Batch'] == batch])  # FIXED: total items in batch

                        if not batch_optical.empty:
                            subfield_counts = batch_optical['Subfield'].value_counts()

                            for subfield, count in subfield_counts.items():
                                percentage = (count / batch_total * 100)  # FIXED: percentage of total items
                                batch_optical_data.append({
                                    'Batch': batch,
                                    'Subfield': subfield,
                                    'Percentage': percentage,
                                    'Count': count
                                })

                    if batch_optical_data:
                        optical_df = pd.DataFrame(batch_optical_data)

                        # Get all unique subfields for consistent coloring
                        all_subfields = optical_df['Subfield'].unique()
                        colors = px.colors.qualitative.Set2[:len(all_subfields)]
                        color_map = {subfield: colors[i] for i, subfield in enumerate(all_subfields)}

                        # Create stacked bar chart with one bar per batch
                        fig_optical_batch = go.Figure()

                        for subfield in all_subfields:
                            subfield_data = optical_df[optical_df['Subfield'] == subfield]

                            # Create lists for batches and percentages
                            batch_list = []
                            percentage_list = []
                            text_list = []
                            count_list = []

                            for batch in selected_batches:
                                batch_subfield_data = subfield_data[subfield_data['Batch'] == batch]
                                if not batch_subfield_data.empty:
                                    percentage = batch_subfield_data['Percentage'].iloc[0]
                                    count = int(batch_subfield_data['Count'].iloc[0])
                                    batch_list.append(batch)
                                    percentage_list.append(percentage)
                                    text_list.append(f'{percentage:.1f}% ({count})')
                                    count_list.append(count)
                                else:
                                    batch_list.append(batch)
                                    percentage_list.append(0)
                                    text_list.append('')
                                    count_list.append(0)

                            fig_optical_batch.add_trace(go.Bar(
                                name=subfield,
                                x=batch_list,
                                y=percentage_list,
                                marker_color=color_map[subfield],
                                text=text_list,
                                textposition='inside',
                                textfont=dict(color='black', size=12),
                                customdata=count_list,
                                hovertemplate=f'<b>{subfield}</b><br>Percentage: %{{y:.1f}}%<br>Count: %{{customdata}}<extra></extra>'
                            ))

                        fig_optical_batch.update_layout(
                            title='Optical Quality Rejection by Batch (% of Total Items)',
                            barmode='stack',
                            xaxis_title='Batch',
                            yaxis_title='Percentage (%)',
                            legend_title='Optical Quality Subfield',
                            showlegend=True
                        )

                        st.plotly_chart(fig_optical_batch, use_container_width=True)

                        # Show optical quality rejection summary by batch
                        optical_summary = optical_rejected_by_batch.groupby('Batch').size().reset_index(
                            name='Optical_Rejections')
                        total_rejections_by_batch = filtered_df[filtered_df['Status'] == 'Rejected'].groupby(
                            'Batch').size().reset_index(name='Total_Rejections')
                        summary_merged = pd.merge(optical_summary, total_rejections_by_batch, on='Batch', how='left')
                        summary_merged['Optical_Percentage'] = (summary_merged['Optical_Rejections'] / summary_merged[
                            'Total_Rejections'] * 100).round(1)

                        st.write("**Optical Quality Rejections Summary by Batch:**")
                        for _, row in summary_merged.iterrows():
                            st.write(
                                f"• {row['Batch']}: {row['Optical_Rejections']} optical quality rejections out of {row['Total_Rejections']} total rejections ({row['Optical_Percentage']:.1f}%)")
                            # NEW: View detailed data option with batch source
                            with st.expander("📋 View Optical Quality Subfield Data by Batch"):
                                # Show optical quality subfields by batch
                                optical_subfield_summary = optical_rejected_by_batch.groupby(
                                    ['Batch', 'Subfield']).size().reset_index(name='Count')

                                # Add batch totals for percentage calculation
                                batch_totals = filtered_df.groupby('Batch').size().reset_index(name='Total')
                                optical_subfield_summary = optical_subfield_summary.merge(batch_totals, on='Batch')
                                optical_subfield_summary['Percentage of Total'] = (
                                            optical_subfield_summary['Count'] / optical_subfield_summary[
                                        'Total'] * 100).round(2)

                                # Display
                                display_cols = ['Batch', 'Subfield', 'Count', 'Percentage of Total']
                                st.dataframe(optical_subfield_summary[display_cols], use_container_width=True)

                                # Also show detailed individual records
                                st.write("**Individual Records:**")
                                st.dataframe(optical_rejected_by_batch[['Batch', 'AIOL_Serial_Number', 'Subfield']],
                                             use_container_width=True)

                                csv_optical_batch = optical_rejected_by_batch.to_csv(index=False).encode('utf-8')
                                st.download_button(
                                    "📥 Download Optical Quality Details CSV",
                                    csv_optical_batch,
                                    "optical_quality_batch_details.csv",
                                    "text/csv",
                                    key='download-optical-quality-batch'
                                )
                        else:
                            st.info("No optical quality rejection data available for selected batches")
                    else:
                        st.info("No optical quality rejections found in selected data")


            # Data table
            with st.expander("📋 View Raw Data"):
                st.dataframe(filtered_df, use_container_width=True)

                # Download option
                csv = filtered_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "Download Filtered Data as CSV",
                    csv,
                    "aiol_rejection_analysis.csv",
                    "text/csv",
                    key='download-csv'
                )

        else:
            st.warning("No data available for the selected filters.")

    else:
        st.error("No valid data found in uploaded files.")

else:
    # Instructions
    st.info("Please upload one or more Excel files containing AIOL production data.")
    st.markdown("""
    ### Expected file format:
    - **Batch information**: Located in the first row (D1:R1), containing text like "AIOL production summary SNXX-SNXX"
    - **Status data**: Located in column A starting from row 6, with format:
  - `Approved` - for approved items
  - `Rejected - Assembly` - for assembly rejections with subfield
  - `Rejected - Optical quality ` - for optical quality rejections with subfield
  - `Rejected - Injection failure` - for injection failure rejections
  - `Rejected - Human error` - for human error rejections
  - `Rejected - Not sealed` - for sealing rejections
  - `Rejected - Failed accommodation` - for accommodation failures
  - `Rejected - Other` - for other rejections
  - `In process` - items still being processed (will be ignored)

    ### Features:
    - Upload multiple Excel files (each representing a different batch)
    - Select which batches to analyze
    - Choose between combined view or side-by-side comparison
    - View rejection reasons statistics with pie charts and bar charts
    - Download filtered data as CSV
    """)