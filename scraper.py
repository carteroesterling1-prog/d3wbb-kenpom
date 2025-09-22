import pandas as pd
import requests
from bs4 import BeautifulSoup
import gspread
from google.oauth2.service_account import Credentials

# ---------- CONFIG ----------
EXCEL_FILE = "2022_NCAA_D3_WBB.xlsx"
SHEET_NAME = "D3WBB Metrics"
WORKSHEET = "2025 Season"

# ---------- GOOGLE SHEETS AUTH ----------
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
gc = gspread.authorize(creds)

sh = gc.open(SHEET_NAME)
try:
    worksheet = sh.worksheet(WORKSHEET)
except:
    worksheet = sh.add_worksheet(title=WORKSHEET, rows="2000", cols="30")

# ---------- SCRAPE TEAM STATS ----------
def get_team_stats(team_id):
    url = f"https://stats.ncaa.org/teams/{team_id}"
    r = requests.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    stats = {}
    try:
        table = soup.find("table", {"class": "my_stats"})
        rows = table.find_all("tr")[1:]
        for row in rows:
            cols = [c.get_text(strip=True) for c in row.find_all("td")]
            if len(cols) >= 2:
                stats[cols[0]] = cols[1:]
    except Exception as e:
        print(f"Error scraping team {team_id}: {e}")

    return stats

# ---------- METRICS CALC ----------
def compute_metrics(stats):
    # Convert strings to ints safely
    def val(key, idx=0):
        try:
            return int(stats.get(key, ["0"])[idx])
        except:
            return 0

    FGM = val("FGM")
    FGA = val("FGA")
    TPM = val("3FGM")
    FTA = val("FTA")
    TO = val("TO")
    OREB = val("Offensive Rebounds")
    OppDREB = val("Opponent Defensive Rebounds")
    Points = val("PTS")

    OppFGM = val("Opp FGM")
    OppFGA = val("Opp FGA")
    OppTPM = val("Opp 3FGM")
    OppFTA = val("Opp FTA")
    OppTO = val("Opp TO")
    OppOREB = val("Opponent Offensive Rebounds")
    DREB = val("Defensive Rebounds")
    OppPoints = val("Opp PTS")

    # Possessions
    poss = FGA - OREB + TO + (0.475 * FTA)
    opp_poss = OppFGA - OppOREB + OppTO + (0.475 * OppFTA)

    # Efficiencies
    AdjO = (Points / poss * 100) if poss > 0 else 0
    AdjD = (OppPoints / opp_poss * 100) if opp_poss > 0 else 0
    Tempo = (poss + opp_poss) / 2  # per game, needs dividing by GP if aggregated

    # Four Factors
    eFG = ((FGM + 0.5 * TPM) / FGA) if FGA > 0 else 0
    TOVp = (TO / poss) if poss > 0 else 0
    OREBp = (OREB / (OREB + OppDREB)) if (OREB + OppDREB) > 0 else 0
    FTRate = (FTA / FGA) if FGA > 0 else 0

    return {
        "Poss": round(poss, 1),
        "OppPoss": round(opp_poss, 1),
        "AdjO": round(AdjO, 2),
        "AdjD": round(AdjD, 2),
        "Tempo": round(Tempo, 1),
        "eFG%": round(eFG, 3),
        "TOV%": round(TOVp, 3),
        "OREB%": round(OREBp, 3),
        "FTRate": round(FTRate, 3),
    }

# ---------- MAIN ----------
def main():
    team_ids = pd.read_excel(EXCEL_FILE)

    all_data = []
    for _, row in team_ids.iterrows():
        team_id = row["TEAM_ID"]
        team_name = row.get("TEAM_NAME", "Unknown")

        try:
            raw_stats = get_team_stats(team_id)
            metrics = compute_metrics(raw_stats)
            metrics["Team"] = team_name
            all_data.append(metrics)
            print(f"Processed {team_name}")
        except Exception as e:
            print(f"Error with {team_name}: {e}")

    df = pd.DataFrame(all_data)

    # Reorder columns
    cols = ["Team", "AdjO", "AdjD", "Tempo", "eFG%", "TOV%", "OREB%", "FTRate", "Poss", "OppPoss"]
    df = df[cols]

    worksheet.clear()
    worksheet.update([df.columns.values.tolist()] + df.values.tolist())

if __name__ == "__main__":
    main()
