import streamlit as st
import pandas as pd
from io import BytesIO
import math
import base64
from pathlib import Path

st.set_page_config(page_title="MCG Suite Clearing List Generator", layout="wide")

def image_to_base64(image_path):
    path = Path(image_path)
    if not path.exists():
        return ""
    return base64.b64encode(path.read_bytes()).decode()


mcg_image_base64 = image_to_base64("MCG_Pic.jpeg")

page_style = """
<style>
.stApp {
    background:
        linear-gradient(180deg, rgba(8, 18, 32, 0.92), rgba(13, 24, 38, 0.96)),
        #0b1624;
    color: #f8fafc;
}

[data-testid="stHeader"] {
    background: rgba(8, 18, 32, 0.78);
}

.block-container {
    padding-top: 1.4rem;
    max-width: 1250px;
}

.mcg-hero {
    min-height: 310px;
    border-radius: 18px;
    background-image:
        linear-gradient(90deg, rgba(5, 12, 22, 0.88), rgba(5, 12, 22, 0.42), rgba(5, 12, 22, 0.10)),
        url("data:image/jpeg;base64,IMAGE_HERE");
    background-size: cover;
    background-position: center;
    display: flex;
    align-items: flex-end;
    padding: 34px;
    margin-bottom: 28px;
    border: 1px solid rgba(255, 255, 255, 0.16);
    box-shadow: 0 18px 45px rgba(0, 0, 0, 0.34);
}

.mcg-hero-content {
    max-width: 760px;
}

.mcg-kicker {
    color: #fbbf24;
    font-size: 0.88rem;
    font-weight: 800;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    margin-bottom: 8px;
}

.mcg-title {
    font-size: 2.65rem;
    line-height: 1.05;
    font-weight: 900;
    color: #ffffff;
    margin-bottom: 10px;
}

.mcg-subtitle {
    color: #dbeafe;
    font-size: 1.08rem;
    max-width: 650px;
}

h2, h3 {
    color: #ffffff !important;
}

div[data-testid="stSubheader"] {
    color: #ffffff;
}

.stButton > button {
    background: linear-gradient(135deg, #f97316, #ef4444);
    color: white;
    border: 0;
    border-radius: 10px;
    padding: 0.65rem 1.05rem;
    font-weight: 800;
    box-shadow: 0 10px 22px rgba(239, 68, 68, 0.28);
}

.stButton > button:hover {
    background: linear-gradient(135deg, #fb923c, #f43f5e);
    color: white;
    border: 0;
}

div[data-testid="stAlert"] {
    border-radius: 12px;
}

div[data-testid="stExpander"] {
    background: rgba(255, 255, 255, 0.045);
    border: 1px solid rgba(255, 255, 255, 0.12);
    border-radius: 12px;
}

.stDataFrame {
    border-radius: 12px;
    overflow: hidden;
    border: 1px solid rgba(255, 255, 255, 0.12);
}

label, p, span {
    color: #e5e7eb;
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #0ea5e9, #2563eb);
    color: white;
    border: 0;
    border-radius: 10px;
    font-weight: 800;
}
</style>
"""

st.markdown(
    page_style.replace("IMAGE_HERE", mcg_image_base64),
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="mcg-hero">
        <div class="mcg-hero-content">
            <div class="mcg-kicker">MCG Operations</div>
            <div class="mcg-title">Suite Runner Clearing List Generator</div>
            <div class="mcg-subtitle">
                Build a clean runner list with balanced suite counts, special room handling,
                and ready-to-download Excel output.
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)




def suite(level, number):
    return f"{level}-{number:02d}"


def suite_range(level, start, end):
    return [suite(level, i) for i in range(start, end + 1)]


AREAS = {
    "Area 1: 2-01 to 2-10": suite_range(2, 1, 10),
    "Area 2: 2-11 and 2-12": suite_range(2, 11, 12),
    "Area 3: 3-01 to 3-10": suite_range(3, 1, 10),
    "Area 4: 3-11 to 3-20": suite_range(3, 11, 20),
    "Area 5: 3-21 to 3-30": suite_range(3, 21, 30),
    "Area 6: 3-31 to 3-44": suite_range(3, 31, 44),
    "Area 7: 3-45 to 3-54": suite_range(3, 45, 54),
    "Media Box Suites: 3-55, 3-56, 3-57": suite_range(3, 55, 57),
    "Area 8: 3-58 to 3-69": suite_range(3, 58, 69),
    "Area 9: 3-70 to 3-81": suite_range(3, 70, 81),
    "Area 10: 3-82 to 3-92": suite_range(3, 82, 92),
    "Area 11: 3-93 to 3-102": suite_range(3, 93, 102),
    "Area 12: 3-103 to 3-110": suite_range(3, 103, 110),
}

SPECIAL_GROUPS = {
    "NISSAN": {"suites": suite_range(3, 14, 16), "low_limit": 28, "medium_limit": 39},
    "Jolimont": {"suites": suite_range(3, 47, 48), "low_limit": 30, "medium_limit": 39},
    "Jameson": {"suites": suite_range(3, 79, 81), "low_limit": 22, "medium_limit": 35},
}

SPECIAL_NAMES = {
    suite_name: name
    for name, group_data in SPECIAL_GROUPS.items()
    for suite_name in group_data["suites"]
}

SPECIAL_PRIMARY_SUITE = {
    "NISSAN": "3-14",
    "Jolimont": "3-47",
    "Jameson": "3-79",
}

SPECIAL_DISPLAY_LABELS = {
    "NISSAN": "3-14-16 (Nissan)",
    "Jolimont": "3-47 (Jolimont)",
    "Jameson": "3-79 (Jameson)",
}

SPECIAL_SUITE_TO_PRIMARY = {
    suite_name: SPECIAL_PRIMARY_SUITE[group_name]
    for group_name, group_data in SPECIAL_GROUPS.items()
    for suite_name in group_data["suites"]
}

DISPLAY_TO_SUITES = {
    SPECIAL_DISPLAY_LABELS[group_name]: group_data["suites"]
    for group_name, group_data in SPECIAL_GROUPS.items()
}

TOGETHER_GROUPS = [
    suite_range(2, 1, 10),
    suite_range(2, 11, 12),
    suite_range(3, 14, 16),
    suite_range(3, 47, 48),
    suite_range(3, 55, 57),
    suite_range(3, 79, 81),
    ["3-58", "3-59"],
    ["3-60", "3-61"],
    ["3-62", "3-63"],
    ["3-64", "3-65"],
    ["3-66", "3-67"],
    ["3-68", "3-69"],
    ["3-70", "3-71"],
    ["3-72", "3-73"],
    ["3-74", "3-75"],
    ["3-76", "3-77"],
    ["3-82", "3-83"],
    ["3-84", "3-85"],
    ["3-86", "3-87"],
    ["3-88", "3-89"],
    ["3-90", "3-91"],
    ["3-93", "3-94"],
    ["3-95", "3-96"],
    ["3-97", "3-98"],
    ["3-99", "3-100"],
    ["3-101", "3-102"],
    ["3-103", "3-104"],
    ["3-105", "3-106"],
    ["3-107", "3-108"],
    ["3-109", "3-110"],
]

ROUTE_ORDER = (
    suite_range(3, 1, 36)
    + suite_range(2, 1, 10)
    
    + suite_range(3, 37, 44)
    + suite_range(3, 45, 54)
    + suite_range(2, 11, 12)
    + suite_range(3, 55, 110)
)

ROUTE_INDEX = {s: i for i, s in enumerate(ROUTE_ORDER)}

CLOSE_TO_KITCHEN = set(suite_range(3, 31, 44))
MIDDLE_AREAS = set(suite_range(3, 21, 57))
LAST_AREAS = set(suite_range(3, 82, 110))

LEVEL_2_01_TO_10 = set(suite_range(2, 1, 10))
LEVEL_2_01_TO_10_TOP_UP = suite_range(3, 17, 36)


def get_all_suites():
    all_suites = []
    for suites in AREAS.values():
        all_suites.extend(suites)
    return all_suites


def sort_by_route(suites):
    return sorted(suites, key=lambda x: ROUTE_INDEX.get(x, 9999))


def display_option_for_suite(suite_name):
    if suite_name not in SPECIAL_SUITE_TO_PRIMARY:
        return suite_name

    primary_suite = SPECIAL_SUITE_TO_PRIMARY[suite_name]
    if suite_name != primary_suite:
        return None

    group_name = SPECIAL_NAMES[suite_name]
    return SPECIAL_DISPLAY_LABELS[group_name]


def display_options_for_area(suites):
    options = []
    for suite_name in suites:
        display_name = display_option_for_suite(suite_name)
        if display_name:
            options.append(display_name)
    return options


def expand_display_options(selected_options):
    suites = []
    for option in selected_options:
        suites.extend(DISPLAY_TO_SUITES.get(option, [option]))
    return suites


def get_runner_labels(runner_count, raw_labels):
    labels = [line.strip() for line in raw_labels.splitlines() if line.strip()]
    return [
        labels[index] if index < len(labels) else f"Runner {index + 1}"
        for index in range(runner_count)
    ]


def pack_group_load(group_name, packs):
    group_data = SPECIAL_GROUPS[group_name]

    if packs <= group_data["low_limit"]:
        return 1
    if packs <= group_data["medium_limit"]:
        return 2
    return 3


def location_multiplier(suite_name):
    if suite_name in LAST_AREAS:
        return 1.15
    if suite_name in CLOSE_TO_KITCHEN:
        return 0.85
    if suite_name in MIDDLE_AREAS:
        return 0.95
    return 1.0


def build_suite_loads(open_suites, pack_data):
    open_set = set(open_suites)
    loads = {s: 1.0 for s in open_suites}

    for group_name, group_data in SPECIAL_GROUPS.items():
        open_in_group = [s for s in group_data["suites"] if s in open_set]
        if not open_in_group:
            continue

        group_load = pack_group_load(group_name, pack_data[group_name])
        share = group_load / len(open_in_group)

        for s in open_in_group:
            loads[s] = share

    return {s: loads[s] * location_multiplier(s) for s in open_suites}


def build_assignment_units(open_suites):
    open_set = set(open_suites)
    used = set()
    units = []

    group_by_first = {}
    for group in TOGETHER_GROUPS:
        selected = sort_by_route([s for s in group if s in open_set])
        if selected:
            group_by_first[selected[0]] = selected

    for s in ROUTE_ORDER:
        if s not in open_set or s in used:
            continue

        unit = group_by_first.get(s, [s])
        units.append(unit)
        used.update(unit)

    return units


def unit_load(unit, suite_loads):
    return sum(suite_loads[s] for s in unit)


def split_units_evenly(units, runner_count, suite_loads):
    if not units:
        return [[] for _ in range(runner_count)]

    active_runner_count = min(runner_count, len(units))
    unit_loads = [unit_load(unit, suite_loads) for unit in units]
    unit_counts = [len(unit) for unit in units]
    total_load = sum(unit_loads)
    total_count = sum(unit_counts)
    target_load = total_load / active_runner_count
    target_count = total_count / active_runner_count

    prefix_loads = [0.0]
    prefix_counts = [0]

    for load, count in zip(unit_loads, unit_counts):
        prefix_loads.append(prefix_loads[-1] + load)
        prefix_counts.append(prefix_counts[-1] + count)

    infinity = float("inf")
    dp = [[infinity] * (len(units) + 1) for _ in range(active_runner_count + 1)]
    split_at = [[0] * (len(units) + 1) for _ in range(active_runner_count + 1)]
    dp[0][0] = 0.0

    for group_count in range(1, active_runner_count + 1):
        for end in range(group_count, len(units) + 1):
            best_cost = infinity
            best_start = group_count - 1

            for start in range(group_count - 1, end):
                previous_cost = dp[group_count - 1][start]
                if previous_cost == infinity:
                    continue

                group_load = prefix_loads[end] - prefix_loads[start]
                group_suite_count = prefix_counts[end] - prefix_counts[start]
                load_cost = (group_load - target_load) ** 2
                count_cost = (group_suite_count - target_count) ** 2 * 0.04
                cost = previous_cost + load_cost + count_cost

                if cost < best_cost:
                    best_cost = cost
                    best_start = start

            dp[group_count][end] = best_cost
            split_at[group_count][end] = best_start

    partitions = []
    end = len(units)

    for group_count in range(active_runner_count, 0, -1):
        start = split_at[group_count][end]
        partitions.append(units[start:end])
        end = start

    partitions.reverse()
    partitions.extend([[] for _ in range(runner_count - active_runner_count)])
    return partitions


def special_summary_label(group_name, pack_data=None):
    if pack_data is None:
        return SPECIAL_DISPLAY_LABELS[group_name]

    suite_count = pack_group_load(group_name, pack_data[group_name])

    if group_name == "NISSAN":
        label = "3-14-16"
    else:
        label = SPECIAL_PRIMARY_SUITE[group_name]

    return f"{label} ({suite_count} suites)"


def collapse_suite_display(suites, pack_data=None):
    collapsed = []
    used_special_suites = set()

    for suite_name in suites:
        if suite_name in used_special_suites:
            continue

        if suite_name in SPECIAL_SUITE_TO_PRIMARY:
            group_name = SPECIAL_NAMES[suite_name]
            group_suites = SPECIAL_GROUPS[group_name]["suites"]
            used_special_suites.update(group_suites)
            collapsed.append(special_summary_label(group_name, pack_data))
        else:
            collapsed.append(suite_name)

    return collapsed


def display_suite_count(suites, pack_data=None):
    count = 0
    used_special_suites = set()

    for suite_name in suites:
        if suite_name in used_special_suites:
            continue

        if suite_name in SPECIAL_SUITE_TO_PRIMARY:
            group_name = SPECIAL_NAMES[suite_name]
            group_suites = SPECIAL_GROUPS[group_name]["suites"]
            used_special_suites.update(group_suites)

            if pack_data is None:
                count += 1
            else:
                count += pack_group_load(group_name, pack_data[group_name])
        else:
            count += 1

    return count


def move_suite_between_assignments(assignments, from_index, to_index, suite_name):
    if suite_name not in assignments[from_index]["Suites"]:
        return False

    assignments[from_index]["Suites"].remove(suite_name)
    assignments[to_index]["Suites"].append(suite_name)
    assignments[from_index]["Suites"] = sort_by_route(assignments[from_index]["Suites"])
    assignments[to_index]["Suites"] = sort_by_route(assignments[to_index]["Suites"])

    return True


def refresh_assignment_totals(assignments, suite_loads, pack_data):
    for assignment in assignments:
        assignment["Count"] = display_suite_count(assignment["Suites"], pack_data)
        assignment["Load"] = sum(
            suite_loads[suite_name] for suite_name in assignment["Suites"]
        )


def rebalance_level_2_01_to_10(assignments, suite_loads, pack_data):
    if not assignments:
        return assignments

    total_display_count = sum(
        display_suite_count(assignment["Suites"], pack_data)
        for assignment in assignments
    )

    target_count = math.ceil(total_display_count / len(assignments))

    level_2_runner_index = None

    for index, assignment in enumerate(assignments):
        if LEVEL_2_01_TO_10.issubset(set(assignment["Suites"])):
            level_2_runner_index = index
            break

    if level_2_runner_index is None:
        refresh_assignment_totals(assignments, suite_loads, pack_data)
        return assignments

    level_2_runner = assignments[level_2_runner_index]

    while display_suite_count(level_2_runner["Suites"], pack_data) < target_count:
        moved = False

        for top_up_suite in LEVEL_2_01_TO_10_TOP_UP:
            if top_up_suite in level_2_runner["Suites"]:
                continue

            donor_index = None

            for index, assignment in enumerate(assignments):
                if index == level_2_runner_index:
                    continue

                if top_up_suite in assignment["Suites"]:
                    donor_index = index
                    break

            if donor_index is None:
                continue

            if move_suite_between_assignments(
                assignments,
                donor_index,
                level_2_runner_index,
                top_up_suite,
            ):
                moved = True
                break

        if not moved:
            break

    refresh_assignment_totals(assignments, suite_loads, pack_data)
    return assignments


def assign_suites(units, runner_count, runner_labels, suite_loads, pack_data):
    partitions = split_units_evenly(units, runner_count, suite_loads)
    assignments = []

    for i, runner_units in enumerate(partitions):
        suites = sort_by_route([s for unit in runner_units for s in unit])

        assignments.append(
            {
                "Runner": runner_labels[i],
                "Suites": suites,
                "Count": display_suite_count(suites, pack_data),
                "Load": sum(unit_load(unit, suite_loads) for unit in runner_units),
            }
        )

    return rebalance_level_2_01_to_10(assignments, suite_loads, pack_data)


def area_for_suite(suite_name):
    for area_name, suites in AREAS.items():
        if suite_name in suites:
            return area_name.split(":")[0]

    return ""


def suite_note(suite_name, pack_data):
    if suite_name not in SPECIAL_NAMES:
        return ""

    area_name = SPECIAL_NAMES[suite_name]
    packs = pack_data.get(area_name, 0)
    room_equiv = pack_group_load(area_name, packs)

    return f"{area_name}: {packs} packs, counts as {room_equiv} suite load"


def create_output_tables(assignments, pack_data, suite_loads):
    summary_rows = []
    detail_rows = []

    for runner in assignments:
        summary_rows.append(
            {
                "Runner": runner["Runner"],
                "Suite Count": runner["Count"],
                "Workload Score": round(runner["Load"], 2),
                "Suites": ", ".join(collapse_suite_display(runner["Suites"], pack_data)),
            }
        )

        used_special_suites = set()

        for s in runner["Suites"]:
            if s in used_special_suites:
                continue

            if s in SPECIAL_SUITE_TO_PRIMARY:
                group_name = SPECIAL_NAMES[s]
                group_suites = SPECIAL_GROUPS[group_name]["suites"]
                runner_group_suites = [
                    suite for suite in group_suites if suite in runner["Suites"]
                ]
                used_special_suites.update(runner_group_suites)
                row_suite_name = SPECIAL_DISPLAY_LABELS[group_name]
                row_load = sum(suite_loads[suite] for suite in runner_group_suites)
            else:
                row_suite_name = s
                row_load = suite_loads[s]

            detail_rows.append(
                {
                    "Runner": runner["Runner"],
                    "Area": area_for_suite(s),
                    "Suite": row_suite_name,
                    "Notes": suite_note(s, pack_data),
                    "Workload Score": round(row_load, 2),
                }
            )

    return pd.DataFrame(summary_rows), pd.DataFrame(detail_rows)


def create_excel(assignments, pack_data, suite_loads):
    summary_df, detail_df = create_output_tables(assignments, pack_data, suite_loads)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Runner Summary")
        detail_df.to_excel(writer, index=False, sheet_name="Clearing List")

    return output.getvalue(), summary_df, detail_df


st.subheader("1. Runner details")

runner_count = st.number_input(
    "How many runners are working today?",
    min_value=1,
    max_value=50,
    value=10,
)

runner_label_text = st.text_area(
    "Runner names or numbers (optional, one per line)",
    height=90,
    placeholder="Runner 1\nRunner 2\nRunner 3",
)

runner_labels = get_runner_labels(int(runner_count), runner_label_text)

st.subheader("2. Suite status")

main_status = st.radio(
    "Choose suite status",
    ["All suites are open", "All suites are closed", "Custom selection"],
    horizontal=True,
)

open_suites = []

if main_status == "All suites are open":
    open_suites = get_all_suites()

elif main_status == "All suites are closed":
    open_suites = []

else:
    for area_name, suites in AREAS.items():
        with st.expander(area_name):
            area_status = st.radio(
                f"Status for {area_name}",
                ["All open", "All closed", "Select manually"],
                key=f"status_{area_name}",
                horizontal=True,
            )

            if area_status == "All open":
                open_suites.extend(suites)

            elif area_status == "Select manually":
                display_options = display_options_for_area(suites)
                selected = st.multiselect(
                    f"Select open suites for {area_name}",
                    display_options,
                    key=f"select_{area_name}",
                )
                open_suites.extend(expand_display_options(selected))

st.subheader("3. Special suite pack numbers")

col1, col2, col3 = st.columns(3)

with col1:
    nissan_packs = st.number_input(
        "NISSAN packs (3-14 to 3-16)",
        min_value=0,
        value=0,
    )

with col2:
    jolimont_packs = st.number_input(
        "Jolimont packs (3-47 and 3-48)",
        min_value=0,
        value=0,
    )

with col3:
    jameson_packs = st.number_input(
        "Jameson packs (3-79 to 3-81)",
        min_value=0,
        value=0,
    )

pack_data = {
    "NISSAN": int(nissan_packs),
    "Jolimont": int(jolimont_packs),
    "Jameson": int(jameson_packs),
}

st.subheader("4. Generate clearing list")

if st.button("Generate list", type="primary"):
    open_suites = sort_by_route(list(set(open_suites)))

    if len(open_suites) == 0:
        st.warning("No suites are open today. No clearing list needed.")

    else:
        suite_loads = build_suite_loads(open_suites, pack_data)
        units = build_assignment_units(open_suites)

        assignments = assign_suites(
            units,
            int(runner_count),
            runner_labels,
            suite_loads,
            pack_data,
        )

        total_weight = sum(suite_loads[s] for s in open_suites)

        st.success(
            f"{len(open_suites)} open suites assigned to {int(runner_count)} runners. "
            f"Total workload score: {total_weight:.2f}."
        )

        summary_df, detail_df = create_output_tables(assignments, pack_data, suite_loads)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        for runner in assignments:
            st.markdown(
                f"### {runner['Runner']} - {runner['Count']} suites "
                f"({runner['Load']:.2f} workload score)"
            )

            runner_rows = detail_df[detail_df["Runner"] == runner["Runner"]]
            st.table(runner_rows[["Suite", "Area", "Notes"]])

        excel_file, _, _ = create_excel(assignments, pack_data, suite_loads)

        st.download_button(
            label="Download Excel clearing list",
            data=excel_file,
            file_name="mcg_suite_clearing_list.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )