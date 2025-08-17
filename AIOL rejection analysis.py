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
    'Human error': '#ffcc99',       # Light orange
    'Not sealed': '#ff99cc',        # Light pink
    'Faild accommodation': '#c2c2f0', # Light purple
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


def extract_batch_info(df):
    """Extract batch information from the first row (D1:R1)"""
    try:
        # Check cells D1 through R1 for batch information
        for col_idx in range(3, 18):  # D=3, R=17 (0-indexed)
            if col_idx < df.shape[1]:
                cell_value = str(df.iloc[0, col_idx])
                if 'AIOL production summary' in cell_value or 'SN' in cell_value:
                    # Extract the SN range (batch number)
                    match = re.search(r'SN\d+-SN?\d+', cell_value)
                    if match:
                        return match.group()

        # If no specific pattern found, return a generic identifier
        return "Unknown Batch"
    except Exception as e:
        st.warning(f"Could not extract batch info: {str(e)}")
        return "Unknown Batch"


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
                    # Expected formats:
                    # "Rejected - Assembly - Out of spec"
                    # "Rejected - Optical quality - Low MTF"
                    # "Rejected- Other"
                    # "Rejected - Injection failure"
                    # etc.

                    # Split by ' - ' or '- ' (handle inconsistent spacing)
                    parts = re.split(r'\s*-\s*', value_str)

                    if len(parts) >= 2:
                        # Remove "Rejected" from the first part
                        status = 'Rejected'

                        if len(parts) == 2:
                            # Format: "Rejected - [main_reason]"
                            rejection_reason = parts[1].strip()
                            subfield = 'No subfield'
                        else:
                            # Format: "Rejected - [main_reason] - [subfield]"
                            rejection_reason = parts[1].strip()
                            subfield = parts[2].strip() if len(parts) >= 3 else 'No subfield'

                        parsed_data.append({
                            'Batch': batch_name,
                            'AIOL_Serial_Number': serial_number,
                            'Status': status,
                            'Rejection_Reason': rejection_reason,
                            'Subfield': subfield,
                            'Raw_Data': value_str
                        })
                    else:
                        # Handle malformed rejected entries
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


def load_and_process_files(files):
    """Load and process all uploaded Excel files"""
    all_data = []
    all_other_statuses = set()
    batch_info = {}

    for file in files:
        try:
            # Read the Excel file
            df = pd.read_excel(file, header=None, engine='openpyxl')

            # Extract batch information
            batch_name = extract_batch_info(df)
            batch_info[file.name] = batch_name

            # Parse the status data
            parsed_data, other_statuses = parse_status_data(df, batch_name)

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
                'Assembly', 'Optical quality', 'Injection failure',
                'Human error', 'Not sealed', 'Faild accommodation', 'Other'
            ]

            # Find rejection reasons that don't match expected ones (case-insensitive)
            unexpected_reasons = set()
            for reason in rejected_data['Rejection_Reason'].unique():
                if pd.notna(reason):
                    reason_lower = str(reason).lower().strip()
                    if not any(expected in reason_lower for expected in expected_reasons):
                        unexpected_reasons.add(str(reason))

            if unexpected_reasons:
                st.warning(f"⚠️ Found unexpected rejection reasons (check for typos): {', '.join(unexpected_reasons)}")

    if all_data:
        # Convert to DataFrame
        df_combined = pd.DataFrame(all_data)

        st.success(f"📊 Total processed records: {len(df_combined)}")

        # Display batch information
        st.write("### Batch Information")
        batch_df = pd.DataFrame(list(batch_info.items()), columns=['File', 'Batch'])
        st.dataframe(batch_df, use_container_width=True)

        # Sidebar filters
        st.sidebar.header("Filter Options")

        # Batch selection
        available_batches = df_combined['Batch'].unique()
        selected_batches = st.sidebar.multiselect(
            "Select Batches to Analyze",
            available_batches,
            default=available_batches
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

                # Assembly subfields analysis
                st.write("### Assembly Rejection Analysis")
                assembly_rejected = filtered_df[(filtered_df['Status'] == 'Rejected') &
                                                (filtered_df['Rejection_Reason'].str.lower() == 'assembly')]

                if not assembly_rejected.empty:
                    # Single stacked bar chart for assembly subfields (percentages of total items)
                    total_items = len(filtered_df)  # FIXED: Back to total items
                    subfield_counts = assembly_rejected['Subfield'].value_counts()
                    subfield_percentages = (subfield_counts / total_items * 100).round(1)

                    # Create a single stacked bar chart for assembly subfields
                    fig_assembly = go.Figure()

                    # Define colors for different subfields
                    colors = px.colors.qualitative.Pastel[:len(subfield_percentages)]

                    # Add each subfield as a segment in the stacked bar
                    for i, (subfield, percentage) in enumerate(subfield_percentages.items()):
                        count = subfield_counts[subfield]
                        fig_assembly.add_trace(go.Bar(
                            name=subfield,
                            x=['Assembly Rejection'],
                            y=[percentage],
                            marker_color=colors[i % len(colors)],
                            text=f' {percentage:.1f}% ({count})',
                            textposition='inside',
                            textfont=dict(color='black', size=12),
                            hovertemplate=f'<b>{subfield}</b><br>Percentage: {percentage:.1f}%<br>Count: {count}<extra></extra>'
                        ))

                    fig_assembly.update_layout(
                        title="Assembly Rejection Distribution (%)",
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

                    st.plotly_chart(fig_assembly, use_container_width=True)

                    # Show assembly rejection summary
                    st.info(
                        f"📊 Assembly rejections: {len(assembly_rejected)} out of {len(filtered_df[filtered_df['Status'] == 'Rejected'])} total rejections")
                else:
                    st.info("No assembly rejections found in selected data")

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

                    # Show optical quality rejection summary
                    st.info(
                        f"📊 Optical quality rejections: {len(optical_rejected)} out of {len(filtered_df[filtered_df['Status'] == 'Rejected'])} total rejections")
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
                    else:
                        st.info("No rejection data available for selected batches")
                else:
                    st.info("No rejected items in selected data")

                # Assembly subfields analysis by batch
                st.write("#### Assembly Rejection Subfields by Batch (%)")

                assembly_rejected_by_batch = filtered_df[(filtered_df['Status'] == 'Rejected') &
                                                         (filtered_df['Rejection_Reason'].str.lower() == 'assembly')]

                if not assembly_rejected_by_batch.empty:
                    # Calculate assembly subfield percentages for each batch
                    batch_assembly_data = []

                    for batch in selected_batches:
                        batch_assembly = assembly_rejected_by_batch[assembly_rejected_by_batch['Batch'] == batch]
                        batch_total = len(filtered_df[filtered_df['Batch'] == batch])  # FIXED: total items in batch

                        if not batch_assembly.empty:
                            subfield_counts = batch_assembly['Subfield'].value_counts()

                            for subfield, count in subfield_counts.items():
                                percentage = (count / batch_total * 100)  # FIXED: percentage of total items
                                batch_assembly_data.append({
                                    'Batch': batch,
                                    'Subfield': subfield,
                                    'Percentage': percentage,
                                    'Count': count
                                })

                    if batch_assembly_data:
                        assembly_df = pd.DataFrame(batch_assembly_data)

                        # Get all unique subfields for consistent coloring
                        all_subfields = assembly_df['Subfield'].unique()
                        colors = px.colors.qualitative.Pastel[:len(all_subfields)]
                        color_map = {subfield: colors[i] for i, subfield in enumerate(all_subfields)}

                        # Create stacked bar chart with one bar per batch
                        fig_assembly_batch = go.Figure()

                        for subfield in all_subfields:
                            subfield_data = assembly_df[assembly_df['Subfield'] == subfield]

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

                            fig_assembly_batch.add_trace(go.Bar(
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

                        fig_assembly_batch.update_layout(
                            title='Assembly Rejection by Batch (% of Total Items)',
                            barmode='stack',
                            xaxis_title='Batch',
                            yaxis_title='Percentage (%)',
                            legend_title='Assembly Subfield',
                            showlegend=True
                        )

                        st.plotly_chart(fig_assembly_batch, use_container_width=True)

                        # Show assembly rejection summary by batch
                        assembly_summary = assembly_rejected_by_batch.groupby('Batch').size().reset_index(
                            name='Assembly_Rejections')
                        total_rejections_by_batch = filtered_df[filtered_df['Status'] == 'Rejected'].groupby(
                            'Batch').size().reset_index(name='Total_Rejections')
                        summary_merged = pd.merge(assembly_summary, total_rejections_by_batch, on='Batch', how='left')
                        summary_merged['Assembly_Percentage'] = (summary_merged['Assembly_Rejections'] / summary_merged[
                            'Total_Rejections'] * 100).round(1)

                        st.write("**Assembly Rejections Summary by Batch:**")
                        for _, row in summary_merged.iterrows():
                            st.write(
                                f"• {row['Batch']}: {row['Assembly_Rejections']} assembly rejections out of {row['Total_Rejections']} total rejections ({row['Assembly_Percentage']:.1f}%)")
                    else:
                        st.info("No assembly rejection data available for selected batches")
                else:
                    st.info("No assembly rejections found in selected data")

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
    - **Batch information**: Located in the first row (D1:R1), containing text like "AIOL production summary SN44-S77"
    - **Status data**: Located in column A starting from row 6, with format:
      - `Approved` - for approved items
      - `Rejected - Assembly - Out of spec` - for assembly rejections with subfield
      - `Rejected - Optical quality - Low MTF` - for optical quality rejections with subfield
      - `Rejected - Other` - for other rejections
      - `Rejected - Injection failure` - for injection failure rejections
      - `Rejected - Human error` - for human error rejections
      - `Rejected - Not sealed` - for sealing rejections
      - `Rejected - Failed accommodation` - for accommodation failures
      - `In process` - items still being processed (will be ignored)

    ### Features:
    - Upload multiple Excel files (each representing a different batch)
    - Select which batches to analyze
    - Choose between combined view or side-by-side comparison
    - View rejection reasons statistics with pie charts and bar charts
    - Download filtered data as CSV
    """)
