import math
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException


def calculate_tickets(amount):
    """
    Calculate raffle tickets using the formula:
      tickets = floor((amount - 30) / 20) + 1
    If amount < 30, return 0; cap at 50.
    """
    if amount < 30:
        return 0
    tickets = math.floor((amount - 30) / 20) + 1
    return min(tickets, 50)


def scrape_team_list(driver, url):
    """
    Loads the main team list page and extracts each team's detail URL.
    Returns a list of dictionaries with {"team_name": "", "team_link": ...}.
    (We will fetch the real team name from the detail page.)
    """
    driver.get(url)

    # Dismiss cookie popup if present
    try:
        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "onetrust-accept-btn-handler"))
        ).click()
    except TimeoutException:
        pass

    # Wait for team links to be present; adjust the selector if necessary.
    team_links = WebDriverWait(driver, 15).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.team-list-name a"))
    )

    teams = []
    for link in team_links:
        team_url = link.get_attribute("href")
        teams.append({"team_name": "", "team_link": team_url})

    return teams


def scrape_team_members(driver, team):
    """
    Navigates to a team's detail page and:
      1. Extracts the actual team name from the header.
      2. Extracts each participant's full name and donation amount from the team member blocks.
         If a block's name is "Team Gifts" (case-insensitive), that donation is stored separately.
    Returns a tuple: (members, team_gifts)
      - members: list of dicts {"full_name": ..., "amount": ..., "team": ...}
      - team_gifts: a float donation value (0.0 if not found)
    """
    driver.get(team["team_link"])

    # STEP 1: Get the real team name
    try:
        team_name_el = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.user-info h2"))
        )
        detail_team_name = team_name_el.text.strip()
    except TimeoutException:
        try:
            team_name_el = driver.find_element(By.CSS_SELECTOR, "h1#personal_header")
            detail_team_name = team_name_el.text.strip()
        except:
            detail_team_name = "Unknown Team"
    team["team_name"] = detail_team_name

    # STEP 2: Wait for the team members block to load.
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "#team-members div.members.row div.item")
            )
        )
    except TimeoutException:
        print(f"Could not find participant blocks for {team['team_name']}")
        return [], 0.0

    member_blocks = driver.find_elements(By.CSS_SELECTOR, "#team-members div.members.row div.item")

    members = []
    team_gifts = 0.0  # Will store the donation from "Team Gifts" (if present)

    for block in member_blocks:
        try:
            name_el = block.find_element(By.CSS_SELECTOR, "div.team-roster-participant-name")
            amount_el = block.find_element(By.CSS_SELECTOR, "div.team-roster-participant-raised")

            name = name_el.text.strip()
            amount_str = amount_el.text.strip().replace("$", "").replace(",", "")
            try:
                amount = float(amount_str)
            except ValueError:
                amount = 0.0

            # Identify the "Team Gifts" block
            if name.lower() == "team gifts":
                team_gifts = amount
            else:
                members.append({
                    "full_name": name,
                    "amount": amount,
                    "team": team["team_name"]
                })
        except Exception as e:
            print(f"Skipping a block in {team['team_name']} due to: {e}")
            continue

    return members, team_gifts


def main():
    base_url = "https://support.cancer.ca/site/TR/RelayForLife/?pg=teamlist&fr_id=30024"

    # Initialize Selenium Chrome driver.
    driver = webdriver.Chrome(service=ChromeService())

    teams = scrape_team_list(driver, base_url)

    participant_data = []  # To store individual participant rows.
    team_data = {}  # Dictionary keyed by team name: {"members": [...], "total": float, "team_gifts": float}

    for t in teams:
        members, team_gifts = scrape_team_members(driver, t)
        # Initialize team data entry
        if t["team_name"] not in team_data:
            team_data[t["team_name"]] = {"members": [], "total": 0.0, "team_gifts": team_gifts}
        else:
            team_data[t["team_name"]]["team_gifts"] = team_gifts

        for m in members:
            participant_data.append({
                "full_name": m["full_name"],
                "team": m["team"],
                "amount": m["amount"]
            })
            team_data[t["team_name"]]["members"].append(m["full_name"])
            team_data[t["team_name"]]["total"] += m["amount"]

        # Add the team gifts donation to the team total.
        team_data[t["team_name"]]["total"] += team_gifts

    # Build the Participants DataFrame.
    # Columns: Full Name, Team, Amount Raised, Attendance, Shirt Size, Raffle Tickets.
    # Raffle Tickets for each participant are calculated using the team's "Team Gifts" donation.
    participants_list = []
    for part in participant_data:
        # Look up the team's team_gifts value.
        raffle_tix = calculate_tickets(part["amount"])
        participants_list.append({
            "Full Name": part["full_name"],
            "Team": part["team"],
            "Amount Raised": part["amount"],
            "Attendance": "",  # Empty column for attendance checkbox.
            "Shirt Size": "",  # Empty column for shirt size.
            "Raffle Tickets": raffle_tix
        })
    df_participants = pd.DataFrame(participants_list)

    # Build the Teams DataFrame.
    # Columns: Team Name, Members, Team Gifts, Total Raised, Raffle Tickets, Camera Collected.
    teams_list = []
    for team_name, data in team_data.items():
        raffle_tix = calculate_tickets(data["team_gifts"])
        members_str = ", ".join(data["members"])
        teams_list.append({
            "Team Name": team_name,
            "Members": members_str,
            "Team Gifts": data["team_gifts"],
            "Total Raised": data["total"],
            "Raffle Tickets": raffle_tix,
            "Camera Collected": ""  # Empty column for camera collected checkbox.
        })
    df_teams = pd.DataFrame(teams_list)

    # Write the DataFrames to an Excel file with two sheets.
    output_filename = "RelayForLife.xlsx"
    with pd.ExcelWriter(output_filename, engine="openpyxl") as writer:
        df_participants.to_excel(writer, sheet_name="Participants", index=False)
        df_teams.to_excel(writer, sheet_name="Teams", index=False)

    print(f"Data successfully written to {output_filename}")
    driver.quit()


if __name__ == "__main__":
    main()
