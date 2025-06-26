import streamlit as st
import pandas as pd
from io import BytesIO

# === Constants ===
MAX_STACK_HEIGHT = 4450  # mm
MAX_STACK_WEIGHT = 75    # kg
MIN_COILS = 4
MAX_COILS = 5

st.set_page_config(page_title="BAF Stack Optimizer", layout="wide")
st.title("🔩 BAF Line Stack Optimizer")
st.caption("📥 Upload Excel File (must have 'Width', 'Grade', 'Weight' columns)")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

# === Function to Generate Downloadable Excel ===
def generate_excel(summary_dict, stacks, waiting):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:

        # --- Summary Sheet ---
        summary_df = pd.DataFrame([summary_dict])
        summary_df.to_excel(writer, index=False, sheet_name='Summary')

        # --- Optimized Stacks Sheet ---
        all_stacks_data = []
        for i, stack in enumerate(stacks, 1):
            for coil in stack['Coils']:
                all_stacks_data.append({
                    "Stack No": i,
                    "Stack Grade": stack["Grade"],
                    "Total Stack Width (mm)": stack["Total Width"],
                    "Total Stack Weight (kg)": round(stack["Total Weight"], 2),
                    "Coil Width (mm)": coil["Width"],
                    "Coil Weight (kg)": coil["Weight"],
                    "Coil Grade": coil["Grade"]
                })
        stack_df = pd.DataFrame(all_stacks_data)
        stack_df.to_excel(writer, index=False, sheet_name='Optimized Stacks')

        # --- Waiting Coils Sheet ---
        waiting_df = pd.DataFrame(waiting)
        waiting_df.to_excel(writer, index=False, sheet_name='Waiting Coils')

    output.seek(0)
    return output

# === Main App Logic ===
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    required_cols = {'Width', 'Grade', 'Weight'}
    if not required_cols.issubset(df.columns):
        st.error("❌ Excel must contain columns: 'Width', 'Grade', and 'Weight'")
    else:
        st.success("✅ File uploaded successfully!")

        df['Width'] = pd.to_numeric(df['Width'], errors='coerce')
        df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce')
        df.dropna(subset=['Width', 'Weight'], inplace=True)

        df['Original Grade'] = df['Grade']
        df['Normalized Grade'] = df['Grade'].replace({
            'DR-08': 'T-57',
            'TS-480': 'T-57',
            'DR-75': 'T-57'
        })

        total_input_coils = len(df)
        stacks = []
        waiting = []

        # Stats counters
        stack_4_count = 0
        stack_5_count = 0
        stack_lt_4000 = 0
        stack_ge_4000 = 0

        # Group by normalized grade
        for grade, group in df.groupby("Normalized Grade"):
            group = group.sort_values(by="Width", ascending=False).reset_index(drop=True)
            used = [False] * len(group)

            while True:
                stack = []
                total_width = 0
                total_weight = 0
                stack_indices = []

                for i in range(len(group)):
                    if not used[i] and len(stack) < MAX_COILS:
                        coil_width = group.loc[i, "Width"]
                        coil_weight = group.loc[i, "Weight"]
                        coil_original_grade = group.loc[i, "Original Grade"]

                        if (total_width + coil_width <= MAX_STACK_HEIGHT and
                            total_weight + coil_weight <= MAX_STACK_WEIGHT):
                            stack.append({
                                "Width": coil_width,
                                "Weight": coil_weight,
                                "Grade": coil_original_grade
                            })
                            total_width += coil_width
                            total_weight += coil_weight
                            stack_indices.append(i)

                if len(stack) >= MIN_COILS:
                    for idx in stack_indices:
                        used[idx] = True

                    stacks.append({
                        "Grade": grade,
                        "Total Width": total_width,
                        "Total Weight": total_weight,
                        "Coils": stack
                    })

                    if len(stack) == 4:
                        stack_4_count += 1
                    elif len(stack) == 5:
                        stack_5_count += 1

                    if total_width < 4000:
                        stack_lt_4000 += 1
                    else:
                        stack_ge_4000 += 1
                else:
                    break

            # Add remaining unused coils to waiting list
            for i in range(len(group)):
                if not used[i]:
                    waiting.append({
                        "Grade": group.loc[i, "Original Grade"],
                        "Width": group.loc[i, "Width"],
                        "Weight": group.loc[i, "Weight"]
                    })

        # === Display Results ===
        st.header("📊 Summary")
        summary_dict = {
            "Total Input Coils": total_input_coils,
            "Total Stacks": len(stacks),
            "4-Coil Stacks": stack_4_count,
            "5-Coil Stacks": stack_5_count,
            "Stacks < 4000 mm": stack_lt_4000,
            "Stacks ≥ 4000 mm": stack_ge_4000,
        }

        if stacks:
            avg_stack_height = sum(stack['Total Width'] for stack in stacks) / len(stacks)
            avg_stack_weight = sum(stack['Total Weight'] for stack in stacks) / len(stacks)
            summary_dict["Avg Stack Height (mm)"] = round(avg_stack_height, 2)
            summary_dict["Avg Stack Weight (kg)"] = round(avg_stack_weight, 2)

        st.dataframe(pd.DataFrame([summary_dict]), use_container_width=True)

        # === Optimized Stacks ===
        st.header("📦 Optimized Stacks")
        for i, stack in enumerate(stacks, 1):
            st.subheader(f"Stack {i}: Grade {stack['Grade']}, Width: {stack['Total Width']} mm, Weight: {round(stack['Total Weight'], 2)} kg")
            stack_df = pd.DataFrame(stack['Coils']).reset_index(drop=True)
            stack_df.index += 1
            st.dataframe(stack_df, use_container_width=True)

        # === Waiting Coils ===
        st.header("⏳ Waiting Coils")
        if waiting:
            waiting_df = pd.DataFrame(waiting).reset_index(drop=True)
            waiting_df.index += 1
            st.dataframe(waiting_df, use_container_width=True)
        else:
            st.success("✅ All coils used in valid stacks!")

        st.write(f"Waiting Coils: **{len(waiting)}**")

        # === Download Button ===
        st.download_button(
            label="📥 Download Full Report (Excel)",
            data=generate_excel(summary_dict, stacks, waiting),
            file_name="BAF_Stack_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
