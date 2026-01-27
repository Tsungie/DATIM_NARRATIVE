import pandas as pd
import requests
import time
import os
from datetime import datetime, timedelta
from typing import Dict, List, Any, Tuple, Optional
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# -----------------------------
# 1) SETTINGS
# -----------------------------
BASE_URL = "https://impilo-dhis.mohcc.gov.zw"
USERNAME = os.getenv("DHIS2_USERNAME", "tsncube")
PASSWORD = os.getenv("DHIS2_PASSWORD", "Tsncube@2025")

PROGRAMS = {
    "HTS Register Mobile App": "uwq1KhUSzZ1",
    "ART Register": "ljMn1YcJy83",
}

# Date settings
USER_START_DATE_STR = "01102025"
USER_END_DATE_STR = "31122025"  # Comment out/change logic inside main if needed

# Reliability / speed knobs
PAGE_SIZE = 1000
REQUEST_TIMEOUT = 120
INITIAL_CHUNK_DAYS = 7
MIN_CHUNK_DAYS = 1
CHUNK_RETRIES = 3


# -----------------------------
# 2) HELPER FUNCTIONS
# -----------------------------
def parse_ddmmyyyy(date_str: str) -> datetime:
    return datetime.strptime(date_str, "%d%m%Y")


def make_session(username: str, password: str) -> requests.Session:
    session = requests.Session()
    session.auth = (username, password)

    retry = Retry(
        total=10,
        connect=10,
        read=10,
        status=10,
        backoff_factor=1.0,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        raise_on_status=False,
        respect_retry_after_header=True,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    session.mount("https://", adapter)
    session.mount("http://", adapter)
    session.headers.update({"Accept": "application/json"})
    return session


def dhis2_get_json(session: requests.Session, endpoint: str, params: Dict[str, Any]) -> Dict[str, Any]:
    url = f"{BASE_URL}{endpoint}"
    r = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    return r.json()


def test_connection(session: requests.Session) -> None:
    me = dhis2_get_json(session, "/api/me.json", {"fields": "id,name,username"})
    print(f"✅ Successfully connected to Impilo DHIS2 as: {me.get('name')} ({me.get('username')})")


def build_dataelement_name_map(session: requests.Session, program_uid: str) -> Tuple[Dict[str, str], Dict[str, Optional[str]]]:
    meta = dhis2_get_json(
        session,
        f"/api/programs/{program_uid}.json",
        {
            "fields": (
                "id,name,"
                "programStages[id,programStageDataElements[dataElement[id,name,code,optionSet[id]]]]"
            )
        },
    )

    de_map: Dict[str, str] = {}
    de_optionset_map: Dict[str, Optional[str]] = {}

    for ps in meta.get("programStages", []) or []:
        for psde in ps.get("programStageDataElements", []) or []:
            de = psde.get("dataElement", {}) or {}
            de_id = de.get("id")
            de_name = de.get("name")

            if de_id and de_name:
                de_map[de_id] = de_name

            os_obj = de.get("optionSet")
            if de_id:
                de_optionset_map[de_id] = os_obj.get("id") if isinstance(os_obj, dict) and os_obj.get("id") else None

    return de_map, de_optionset_map


def build_options_name_map(session: requests.Session, option_set_ids: List[str]) -> Dict[str, str]:
    option_map: Dict[str, str] = {}
    unique_sets = sorted(set([x for x in option_set_ids if x]))

    for os_id in unique_sets:
        page = dhis2_get_json(
            session,
            "/api/options.json",
            {
                "filter": f"optionSet.id:eq:{os_id}",
                "paging": "false",
                "fields": "id,code,name",
            },
        )

        for opt in page.get("options", []) or []:
            opt_id = opt.get("id")
            opt_code = opt.get("code")
            opt_name = opt.get("name")

            if opt_name:
                if opt_id:
                    option_map[opt_id] = opt_name
                if opt_code:
                    option_map[opt_code] = opt_name

    return option_map


def daterange_chunks(start: datetime, end: datetime, chunk_days: int) -> List[Tuple[datetime, datetime]]:
    chunks = []
    cur = start
    while cur <= end:
        chunk_end = min(cur + timedelta(days=chunk_days - 1), end)
        chunks.append((cur, chunk_end))
        cur = chunk_end + timedelta(days=1)
    return chunks


def _filter_events_to_end_date(events: List[Dict[str, Any]], end_date: datetime) -> List[Dict[str, Any]]:
    end_d = end_date.date()
    kept = []
    for ev in events:
        raw = ev.get("eventDate")
        if not raw:
            continue
        dt = pd.to_datetime(raw, errors="coerce")
        if pd.isna(dt):
            continue
        if dt.date() <= end_d:
            kept.append(ev)
    return kept


def fetch_events_paged(session: requests.Session, program_uid: str, start_date: datetime, end_date: datetime) -> List[Dict[str, Any]]:
    all_events: List[Dict[str, Any]] = []
    page = 1
    api_end_date = end_date + timedelta(days=1)

    while True:
        payload = dhis2_get_json(
            session,
            "/api/events.json",
            {
                "program": program_uid,
                "startDate": start_date.strftime("%Y-%m-%d"),
                "endDate": api_end_date.strftime("%Y-%m-%d"),
                "orgUnitMode": "ACCESSIBLE",
                "page": page,
                "pageSize": PAGE_SIZE,
                "totalPages": "true",
                "fields": (
                    "event,eventDate,programStage,orgUnit,orgUnitName,status,lastUpdated,"
                    "geometry,dataValues[dataElement,value]"
                ),
            },
        )

        events = payload.get("events", [])
        if not events:
            break

        events = _filter_events_to_end_date(events, end_date)
        all_events.extend(events)

        pager = payload.get("pager", {}) or {}
        page_count = pager.get("pageCount")
        if page_count is not None and page >= int(page_count):
            break

        page += 1

    return all_events


def fetch_chunk_with_no_skip(
    session: requests.Session,
    program_uid: str,
    s: datetime,
    e: datetime,
    program_name: str,
    chunk_no: int,
    total_chunks: int,
    chunk_days: int,
) -> List[Dict[str, Any]]:
    chunk_label = f"{program_name} [{chunk_no}/{total_chunks}] {s.date()} to {e.date()}"
    for attempt in range(1, CHUNK_RETRIES + 1):
        try:
            t0 = time.time()
            events = fetch_events_paged(session, program_uid, s, e)
            dt = time.time() - t0
            rate = (len(events) / dt) if dt > 0 else 0
            print(f"✅ {chunk_label} -> {len(events)} events  ({rate:.1f} events/s)")
            return events
        except Exception as ex:
            print(f"⚠️ {chunk_label} attempt {attempt}/{CHUNK_RETRIES} failed: {ex}")

    if chunk_days > MIN_CHUNK_DAYS:
        new_days = max(MIN_CHUNK_DAYS, chunk_days // 2)
        print(f"🔁 Splitting chunk {s.date()} to {e.date()} into smaller chunks ({new_days} days) to avoid timeouts...")

        all_events: List[Dict[str, Any]] = []
        sub_chunks = daterange_chunks(s, e, new_days)
        sub_total = len(sub_chunks)
        for i, (ss, ee) in enumerate(sub_chunks, start=1):
            sub_events = fetch_chunk_with_no_skip(
                session=session,
                program_uid=program_uid,
                s=ss,
                e=ee,
                program_name=program_name,
                chunk_no=i,
                total_chunks=sub_total,
                chunk_days=new_days,
            )
            all_events.extend(sub_events)
        return all_events

    raise RuntimeError(f"❌ Failed permanently even at {MIN_CHUNK_DAYS}-day chunk: {chunk_label}")


def events_to_dataframe(
    events: List[Dict[str, Any]],
    de_map: Dict[str, str],
    de_optionset_map: Dict[str, Optional[str]],
    option_uid_or_code_to_name: Dict[str, str],
) -> pd.DataFrame:
    records = []
    for event in events:
        row = {
            "Event": event.get("event"),
            "Program stage": event.get("programStage"),
            "Event date": event.get("eventDate"),
            "Organisation unit name": event.get("orgUnitName"),
            "Organisation unit uid": event.get("orgUnit"),
            "Program status": event.get("status"),
            "Last updated on": event.get("lastUpdated"),
            "Longitude": None,
            "Latitude": None,
        }

        geom = event.get("geometry") or {}
        coords = geom.get("coordinates")
        if isinstance(coords, list) and len(coords) >= 2:
            row["Longitude"] = coords[0]
            row["Latitude"] = coords[1]

        for dv in event.get("dataValues", []):
            de_uid = dv.get("dataElement")
            val = dv.get("value")

            if de_uid and de_optionset_map.get(de_uid) and isinstance(val, str):
                translated = option_uid_or_code_to_name.get(val)
                if translated is not None:
                    val = translated

            de_name = de_map.get(de_uid, de_uid)
            row[f"{de_name} [{de_uid}]"] = val

        records.append(row)

    df = pd.DataFrame(records)
    if not df.empty:
        df["Event date"] = pd.to_datetime(df["Event date"], errors="coerce")
        df["Formatted Event date"] = df["Event date"].dt.strftime("%d/%m/%Y")
    return df


# -----------------------------
# 3) MAIN PROCESS
# -----------------------------
def fetch_data() -> pd.DataFrame:
    """
    Fetches data from DHIS2, processes HTS and ART indicators,
    and returns the final aggregated DataFrame by Organisation Unit.
    """
    start_date = parse_ddmmyyyy(USER_START_DATE_STR)
    # Default to USER_END_DATE_STR, else today
    try:
        end_date = parse_ddmmyyyy(USER_END_DATE_STR)
    except NameError:
        end_date = datetime.today()

    if end_date < start_date:
        raise ValueError("End date cannot be earlier than start date.")

    session = make_session(USERNAME, PASSWORD)
    test_connection(session)
    print("➡️ Now fetching event data from DHIS2...")

    detailed_data: Dict[str, pd.DataFrame] = {}

    # --- 1. DOWNLOAD LOOP ---
    for program_name, program_uid in PROGRAMS.items():

        de_map, de_optionset_map = build_dataelement_name_map(session, program_uid)
        
        option_set_ids = [os_id for os_id in de_optionset_map.values() if os_id]
        option_uid_or_code_to_name = build_options_name_map(session, option_set_ids)

        chunk_days = INITIAL_CHUNK_DAYS
        chunks = daterange_chunks(start_date, end_date, chunk_days)
        total_chunks = len(chunks)

        all_events_program: List[Dict[str, Any]] = []

        for idx, (s, e) in enumerate(chunks, start=1):
            pct = int((idx - 1) / total_chunks * 100)
            print(f"🔄 {program_name} progress: {pct}% (starting chunk {idx}/{total_chunks})")

            events = fetch_chunk_with_no_skip(
                session=session,
                program_uid=program_uid,
                s=s,
                e=e,
                program_name=program_name,
                chunk_no=idx,
                total_chunks=total_chunks,
                chunk_days=chunk_days,
            )
            all_events_program.extend(events)

        detailed_data[program_name] = events_to_dataframe(
            all_events_program,
            de_map,
            de_optionset_map,
            option_uid_or_code_to_name
        )

    print("\n✅ ALL PROGRAMS COMPLETED.\n")
    
    HTS = detailed_data.get("HTS Register Mobile App", pd.DataFrame())
    ART = detailed_data.get("ART Register", pd.DataFrame())

    if HTS.empty or ART.empty:
        print("⚠️ Warning: One of the requested programs returned no data.")
        if HTS.empty and ART.empty:
            return pd.DataFrame()

    # Clean column names (remove [UID])
    HTS.columns = HTS.columns.str.replace(r"\s*\[.*?\]", "", regex=True)
    ART.columns = ART.columns.str.replace(r"\s*\[.*?\]", "", regex=True)

    # 2a. HTS_TST (Testing Volume by Facility)
    # Note: "Organisation unit name" is used for grouping
    hts_tst_org = (
        HTS.groupby("Organisation unit name", dropna=False)
        .size()
        .reset_index(name="HTS_TST")
    )
    hts_tst_org.loc["Grand Total"] = ["Grand Total", hts_tst_org["HTS_TST"].sum()]

    # 2b. HTS_POS (Positive Tests by Facility)
    hts_pos_org = (
        HTS.query("`HTS test result` == 'POSITIVE'")
        .groupby("Organisation unit name", dropna=False)
        .size()
        .reset_index(name="HTS_POS")
    )
    hts_pos_org.loc["Grand Total"] = ["Grand Total", hts_pos_org["HTS_POS"].sum()]

    # 2c. TX_NEW (New on ART by Facility)
    tx_new_org = (
        ART.groupby("Organisation unit name", dropna=False)
        .size()
        .reset_index(name="TX_NEW")
    )
    tx_new_org.loc["Grand Total"] = ["Grand Total", tx_new_org["TX_NEW"].sum()]

    # --- 3. MERGE FINAL REPORT (By Facility) ---
    final_org_df = (
        hts_tst_org.query("`Organisation unit name` != 'Grand Total'")
        .merge(hts_pos_org.query("`Organisation unit name` != 'Grand Total'"),
               on="Organisation unit name", how="left")
        .merge(tx_new_org.query("`Organisation unit name` != 'Grand Total'"),
               on="Organisation unit name", how="left")
        .fillna(0)
    )

    # Ensure integer columns
    num_cols = ["HTS_TST", "HTS_POS", "TX_NEW"]
    for col in num_cols:
        if col in final_org_df.columns:
            final_org_df[col] = final_org_df[col].astype(int)

    # Optional: You can also generate the date-based df here if needed,
    # but the prompt asked for "the last dataframe" which in the notebook's 
    # execution flow (highest cell number) was the facility-based one.

    df_mapping = pd.read_csv("datim_mapping_file.csv")
    final_org_df.rename(columns={"Organisation unit name": "Facility"}, inplace=True)

    df_merge = pd.merge(final_org_df, df_mapping, on="Facility", how="left")


    df_merge.to_csv("dhis2_1.csv")

    return df_merge


   