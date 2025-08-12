import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import re
from datetime import datetime

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
                    match = re.search(r'SN\d+-SN\d+', cell_value)
                    if match:
                        return match.group()

        # If no specific pattern found, return a generic identifier
        return "Unknown Batch"
    except Exception as e:
        st.warning(f"Could not extract batch info: {str(e)}")
        return "Unknown Batch"


def parse_status_data(df, batch_name):
    """Parse the status data from column A starting from row 6"""
    try:
        # Get data from column A starting from row 6 (index 5)
        status_data = df.iloc[5:, 0].dropna()  # Column A, starting from row 6

        parsed_data = []
        other_statuses = set()  # Track unexpected statuses

        for idx, value in status_data.items():
            value_str = str(value).strip()

            if value_str and value_str != 'nan':
                # Split by dash
                parts = [part.strip() for part in value_str.split('-')]

                if len(parts) >= 1:
                    status = parts[0].lower()

                    # Handle different status cases
                    if status in ['rejected', 'approved']:
                        rejection_reason = parts[1] if len(parts) >= 2 else 'Unknown'
                        subfield = parts[2] if len(parts) >= 3 else 'No subfield'

                        parsed_data.append({
                            'Batch': batch_name,
                            'Status': status.title(),
                            'Rejection_Reason': rejection_reason,
                            'Subfield': subfield,
                            'Raw_Data': value_str
                        })

                    elif status == 'in process':
                        # Don't count in process items, but track them
                        continue

                    else:
                        # Track other unexpected statuses
                        other_statuses.add(status)

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
            # Define expected rejection reasons
            expected_reasons = ['optical quality', 'assembly', 'injection', 'other', 'sealing', 'in process']

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

                # Combined pie chart
                col1, col2 = st.columns(2)

                with col1:
                    # Pie chart for overall status
                    status_counts = filtered_df['Status'].value_counts()

                    fig_pie = px.pie(
                        values=status_counts.values,
                        names=status_counts.index,
                        title="Overall Status Distribution",
                        color_discrete_sequence=['#0099CC', '#E6F7FF']
                    )
                    fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                    st.plotly_chart(fig_pie, use_container_width=True)

                with col2:
                    # Single stacked bar chart for rejection reasons (percentages)
                    rejected_df = filtered_df[filtered_df['Status'] == 'Rejected']
                    total_rejections = len(filtered_df)
                    if not rejected_df.empty:
                        rejection_counts = rejected_df['Rejection_Reason'].value_counts()
                        rejection_percentages = (rejection_counts / total_rejections * 100).round(1)

                        # Create a single stacked bar chart
                        fig_bar = go.Figure()

                        # Define colors for different rejection reasons
                        colors = px.colors.qualitative.Set3[:len(rejection_percentages)]

                        # Add each rejection reason as a segment in the stacked bar
                        for i, (reason, percentage) in enumerate(rejection_percentages.items()):
                            count = rejection_counts[reason]
                            fig_bar.add_trace(go.Bar(
                                name=reason,
                                x=['Rejection Reasons'],
                                y=[percentage],
                                marker_color=colors[i % len(colors)],
                                text=f'{percentage:.1f}% ({count})',
                                textposition='inside',
                                textfont=dict(color='black', size=12),
                                hovertemplate = f'<b>{reason}</b><br>Percentage: {percentage:.1f}%<br>Count: {count}<extra></extra>'
                            ))

                        fig_bar.update_layout(
                            title="Rejection Reasons Distribution (%)",
                            barmode='stack',
                            #plot_bgcolor='white',
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
                    # Single stacked bar chart for assembly subfields (percentages)
                    total_rejections = len(filtered_df[filtered_df['Status'] == 'Rejected'])
                    subfield_counts = assembly_rejected['Subfield'].value_counts()
                    subfield_percentages = (subfield_counts /total_rejections * 100).round(1)

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
                            hovertemplate = f'<b>{reason}</b><br>Percentage: {percentage:.1f}%<br>Count: {count}<extra></extra>'
                        ))

                    fig_assembly.update_layout(
                        title="Assembly Rejection Distribution (%)",
                        barmode='stack',
                        #plot_bgcolor='white',
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
                    # Single stacked bar chart for optical quality subfields (percentages)
                    total_rejections = len(filtered_df[filtered_df['Status'] == 'Rejected'])
                    subfield_counts = optical_rejected['Subfield'].value_counts()
                    subfield_percentages = (subfield_counts / total_rejections * 100).round(1)

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
                            hovertemplate = f'<b>{reason}</b><br>Percentage: {percentage:.1f}%<br>Count: {count}<extra></extra>'
                        ))

                    fig_optical.update_layout(
                        title="Optical Quality Rejection Distribution (%)",
                        barmode='stack',
                        #plot_bgcolor='white',
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

                # Separate pie charts for each batch
                st.write("#### Status Distribution by Batch")

                # Calculate number of columns for pie charts
                num_batches = len(selected_batches)
                cols_per_row = 3

                # Create pie charts for each batch
                for i in range(0, num_batches, cols_per_row):
                    cols = st.columns(min(cols_per_row, num_batches - i))

                    for j, col in enumerate(cols):
                        if i + j < num_batches:
                            batch = selected_batches[i + j]
                            batch_data = filtered_df[filtered_df['Batch'] == batch]

                            with col:
                                if not batch_data.empty:
                                    status_counts = batch_data['Status'].value_counts()

                                    fig_pie = px.pie(
                                        values=status_counts.values,
                                        names=status_counts.index,
                                        title=f"{batch}",
                                        color_discrete_sequence=['#0099CC', '#E6F7FF']
                                    )
                                    fig_pie.update_traces(
                                        textposition='inside',
                                        textinfo='percent+label',
                                        textfont_size=10
                                    )
                                    fig_pie.update_layout(height=400)
                                    st.plotly_chart(fig_pie, use_container_width=True)
                                else:
                                    st.info(f"No data for {batch}")

                # Rejection reasons bar chart - single stacked bar for each batch
                st.write("#### Rejection Reasons Distribution by Batch (%)")

                rejected_by_batch = filtered_df[filtered_df['Status'] == 'Rejected']

                if not rejected_by_batch.empty:
                    # Calculate rejection reason percentages for each batch
                    batch_rejection_data = []

                    for batch in selected_batches:
                        batch_rejected = rejected_by_batch[rejected_by_batch['Batch'] == batch]

                        if not batch_rejected.empty:
                            batch_total_rejected = len(filtered_df[filtered_df['Batch'] == batch])
                            reason_counts = batch_rejected['Rejection_Reason'].value_counts()

                            for reason, count in reason_counts.items():
                                percentage = (count / batch_total_rejected * 100)
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
                        color_map = {reason: colors[i] for i, reason in enumerate(all_reasons)}

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
                                    count= int(batch_reason_data['Count'].iloc[0])
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
                                marker_color=color_map[reason],
                                text=f' {percentage:.1f}% ({count})',
                                textposition='inside',
                                textfont=dict(color='black', size=12),
                                customdata=count_list,
                                hovertemplate=f'<b>{reason}</b><br>Percentage: {percentage:.1f}%<br>Count: {count}<extra></extra>'
                            ))

                        fig_rejection.update_layout(
                            title='Rejection Reasons by Batch (% of Rejected Items)',
                            barmode='stack',
                            #plot_bgcolor='white',
                            xaxis_title='Batch',
                            yaxis_title='Percentage (%)',
                            legend_title='Rejection Reason',
                            showlegend=True
                        )

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

                        if not batch_assembly.empty:
                            batch_total_rejections = len(
                                filtered_df[(filtered_df['Batch'] == batch) & (filtered_df['Status'] == 'Rejected')])
                            subfield_counts = batch_assembly['Subfield'].value_counts()

                            for subfield, count in subfield_counts.items():
                                percentage = (count / batch_total_rejections * 100)
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
                                hovertemplate=f'<b>{reason}</b><br>Percentage: %{{y:.1f}}%<br>Count: %{{customdata}}<extra></extra>'
                            ))

                        fig_assembly_batch.update_layout(
                            title='Assembly Rejection by Batch (% of Assembly Rejections)',
                            barmode='stack',
                            #plot_bgcolor='white',
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

                        if not batch_optical.empty:
                            batch_total_rejections = len(filtered_df[(filtered_df['Batch'] == batch) & (filtered_df['Status'] == 'Rejected')])
                            subfield_counts = batch_optical['Subfield'].value_counts()

                            for subfield, count in subfield_counts.items():
                                percentage = (count / batch_total_rejections * 100)
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
                                hovertemplate=f'<b>{reason}</b><br>Percentage: %{{y:.1f}}%<br>Count: %{{customdata}}<extra></extra>'
                            ))

                        fig_optical_batch.update_layout(
                            title='Optical Quality Rejection by Batch (% of Optical Quality Rejections)',
                            barmode='stack',
                            #plot_bgcolor='white',
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
      - `Rejected - [reason]` - for rejected items with reason
      - `Rejected - [reason] - [subfield]` - for rejected items with reason and subfield
      - `in process` - items still being processed (will be ignored)

    ### Features:
    - Upload multiple Excel files (each representing a different batch)
    - Select which batches to analyze
    - Choose between combined view or side-by-side comparison
    - View rejection reasons statistics with pie charts and bar charts
    - Download filtered data as CSV

    """)
