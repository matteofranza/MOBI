"""
@author: mfnz
"""
# -------------------------------------------------------------------------
#%% LIBRARIES & GENERAL SETTINGS
# -------------------------------------------------------------------------
import streamlit as st #check streamlit version with st.__version__ 
import pandas as pd
import numpy as np
import re
from datetime import date
import os
from PIL import Image
import base64
from pathlib import Path
import io
from io import BytesIO
import xlsxwriter
import plotly.express as px



UEFA_logo_path = "https://intellcentreinsights-uefa-com.go-vip.net/wp-content/uploads/2023/03/uefa-logo.png"

# set_page_config must be the first st call in order to work
st.set_page_config(page_title="MOBI - Traveling Demand Simulation", page_icon=UEFA_logo_path, layout="wide")

#css
with open("style.css") as f:
    css = f.read()
# Embed the CSS in the Streamlit app
st.markdown(f'<style>{css}</style>', unsafe_allow_html=True)

#Hide "made with streamlit" at the bottom of the page
hide_streamlit_style = """
            <style>
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 

#fix the space before the title
st.markdown("""
        <style>
               .block-container {
                    padding-top: 0rem;
                }
        </style>
        """, unsafe_allow_html=True)

#remove the view full screen option from images
hide_img_fs = '''
<style>
button[title="View fullscreen"]{
    visibility: hidden;}
</style>
'''

st.markdown(hide_img_fs, unsafe_allow_html=True)

#%%
# ------------------------------------------------------------------------------------------------------------------------------------------------------------
#%% TITLE
# ------------------------------------------------------------------------------------------------------------------------------------------------------------
with st.container():
    st.markdown("""
        <div class="centered-container">
            <h1 class="centered-header">MOBI - Traveling Demand Simulation</h1>
        </div>
        """, unsafe_allow_html=True)

#%%
# ------------------------------------------------------------------------------------------------------------------------------------------------------------
#%% LOAD DATA
# ------------------------------------------------------------------------------------------------------------------------------------------------------------

df_raw = pd.read_excel('SF_MOBISimulationTool.xlsx')
#test only on these data
df = df_raw[(df_raw['Season'] == 2023) & (df_raw['Competition'] == "UEL") & (df_raw['RoundType'] == "R16")]

#remove all the contextual info

df = df.drop(columns=['Matchday',
                      'Session',
                      'Date',
                      'Kick-Off CET Time',
                      'Kick-Off Local Time',
                      'Aggregate',
                      'Group',
                      'EndedAfterPenaltyShootout',
                      'Home PenaltyShootout',
                      'Away PenaltyShootout',
                      'Result',
                      'Home Goals',
                      'Away Goals',
                      'EndedAfterPenaltyShootout',
                      'Weather',
                      'Temperature',
                      'Pitch Condition',
                      'WindSpeed kmh',                       
                      'Humidity'])

#reshape the df to have only 1 datapoint for each team (not home/away)
df.columns = [col.replace('Home ', '') if col.startswith('Home ') else col for col in df.columns]
df.columns = [col.replace('_home', '') if col.endswith('_home') else col for col in df.columns]

df = df.drop(columns=[col for col in df.columns if col.startswith('Away ')])
df = df.drop(columns=[col for col in df.columns if col.endswith('_away')])

df = df[df['is_teamVenue_hosting_final'] != "yes"]#in case one of the teams is hosting the final

#%%
# ------------------------------------------------------------------------------------------------------------------------------------------------------------
#%% SIDEBAR INPUTS
# ------------------------------------------------------------------------------------------------------------------------------------------------------------
UEFA_logo = f'<img src="{UEFA_logo_path}" class="center">'

is_button_pressed = 0
with st.sidebar:
    st.sidebar.markdown(UEFA_logo, unsafe_allow_html=True)
    st.markdown('<div style="text-align: center; font-weight: bold; font-size:1.5rem ">Intelligence Centre</div>', unsafe_allow_html=True)

    simulation_mode = st.selectbox("**Simulation mode***", ("", "Data Driven", "Free"))
    
    if simulation_mode == "Data Driven":
         st.markdown("""<b>*Data Driven</b>: simulation based on pre-defined data inputs""", unsafe_allow_html=True)
         
         #simulation parameters
         season = st.selectbox('**Season (title year)**', df['Season'].unique())
         competition = st.selectbox('**Competition**', df['Competition'].unique())
         competition_round = st.selectbox('**Round**', df['RoundType'].unique())
         trials = st.selectbox( "**Trials**", (10, 100, 1000, 10000, 100000, 1000000))  
         
         df_simulation = df[(df['Season'] == season) & 
                            (df['Competition'] == competition) & 
                            (df['RoundType'] == competition_round)]       
         
         colbutton1, colbutton2 = st.columns(2) 
         with colbutton1: 
             if st.button('Run Simulation', type="primary"): #the only reason to define this condition is to allow the button in the sidebar without defining a session state
                 is_button_pressed = 1 
             else:
                 is_button_pressed = 0

         with colbutton2:
            date = date.today()
            date = date.strftime("%Y/%m/%d")
            fileName = "mobi_traveling_demand_" + date + ".xlsx"
            
            @st.cache_data
            def to_excel(df_simulation) -> bytes:
                dataframe_displayed_filtered = io.BytesIO()
                writer = pd.ExcelWriter(dataframe_displayed_filtered, engine="xlsxwriter")
                df_simulation.to_excel(writer, index=False, sheet_name="Sheet1")
                writer.close()
                processed_data = dataframe_displayed_filtered.getvalue()
                return processed_data
            st.download_button(
                label="Download Excel",
                data = to_excel(df_simulation),
                file_name = fileName,
                mime='application/vnd.ms-excel',
                key="green_button",
                type="secondary"
                )
            #define download_button color
            st.markdown("""
                        <style>
                        div.row-widget.stDownloadButton > button:first-child {  
                        padding-top: 0rem;
                        background-color: #01723a;
                        color:#ffffff;
                        }
                        </style>""", unsafe_allow_html=True)
            
    if simulation_mode == "Free":
        st.markdown("""<b>*Free</b>: inputs are fully customizable""", unsafe_allow_html=True)  
        trials = st.selectbox( "**Trials**", (10, 100, 1000, 10000, 100000, 1000000)) 
        if st.button('Run Simulation', type="primary"): #the only reason to define this condition is to allow the button in the sidebar without defining a session state
            is_button_pressed = 1 
        else:
            is_button_pressed = 0


#%%
# ------------------------------------------------------------------------------------------------------------------------------------------------------------
#%% CLUBS POOL
# ------------------------------------------------------------------------------------------------------------------------------------------------------------

if simulation_mode == "Data Driven":  
    with st.container():
        st.markdown("""
                    <div class="centered-container">
                    <h4 class="centered-header">Clubs pool</h4>
                    </div>
                    """, unsafe_allow_html=True)
        
    images_per_row = 4
    for start_idx in range(0, len(df_simulation), images_per_row):
        row_data = df_simulation.iloc[start_idx:start_idx + images_per_row]
        cols = st.columns(images_per_row)
        for col, (_, item) in zip(cols, row_data.iterrows()):
            with col:
                col.markdown(
                    f"<div style='text-align: center'>"
                    f"<img src='{item['club_logo_TFM']}' style='width: 50px;'>"
                    f"</div>", 
                    unsafe_allow_html=True)

#%%
# ------------------------------------------------------------------------------------------------------------------------------------------------------------
#%% RUN SIMULATION WITH BUTTON
# ------------------------------------------------------------------------------------------------------------------------------------------------------------

if is_button_pressed == 1 and simulation_mode == "Data Driven":
                
        simulation = []
        matchups = []
        
        np.random.seed(99)
        
        for _ in range(trials):
            finalists = df_simulation.sample(n=2, weights='competition_champion_prob')
            finalists = finalists.sort_values(by='Team')
            team1, team2 = finalists['Team'].values
            travel_demand1, travel_demand2 = finalists['traveling_demand_FSE_MOBI_EXP'].values
            club_logo_TFM1, club_logo_TFM2 = finalists['club_logo_TFM'].values
            total_travel_demand = travel_demand1 + travel_demand2
            
            simulation.append({
                 'Team1': team1,
                 'Team2': team2,
                 'TravelDemand1': travel_demand1,
                 'TravelDemand2': travel_demand2,
                 'club_logo_TFM1': club_logo_TFM1,
                 'club_logo_TFM2': club_logo_TFM2,
                 'TotalTravelDemand': total_travel_demand
                 })
            simulation_results = pd.DataFrame(simulation)
            
            matchup = (f"{team1} vs {team2}", club_logo_TFM1, club_logo_TFM2)
            matchups.append(matchup)
        #calculate the most common matchup
        matchups = pd.DataFrame(matchups, columns = ['match', 'club_logo_TFM1', 'club_logo_TFM2'] )
        matchups_percentages = pd.DataFrame((matchups['match'].value_counts() / trials) * 100).rename(columns={'count': 'percentage'})
        matchups = matchups.drop_duplicates(subset=['match'])
        matchups_percentages = matchups_percentages.merge(matchups[['match', 'club_logo_TFM1', 'club_logo_TFM2']], on='match', how='left')
        top_matchups = matchups_percentages.sort_values(by='percentage', ascending=False).head(5)
        

        # calculate average travel demand    
        average_demand = round(simulation_results['TotalTravelDemand'].mean())
        st.markdown(f"#### **Estimated Average Traveling Demand:** {average_demand}")

        col1, col2 = st.columns(2)  


        with col2:
            st.markdown("""
                        <div class="centered-container">
                        <h4 class="centered-header">Most Frequent Matchups</h4>
                        </div>
                        """, unsafe_allow_html=True)
            
            for index, row in top_matchups.iterrows():
                st.markdown(f"""
                <style>
                .row {{
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }}
                .column {{
                    flex: 1;
                    text-align: center;
                    padding: 3px;
                }}
                .image {{
                    max-width: 50px;  /* Adjust this to control image width */
                    max-height: 50px; /* Adjust this to control image height */
                    width: auto;
                    height: auto;
                }}
                </style>
                <div class="row">
                    <div class="column">
                        <img src="{row['club_logo_TFM1']}" alt="Fig1" class="image">
                    </div>
                    <div class="column">
                        {row['match']}
                    </div>
                    <div class="column">
                        <img src="{row['club_logo_TFM2']}" alt="Fig2" class="image">
                    </div>
                    <div class="column">
                        {round(row['percentage'])}%
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
            #plot
            with col1:
                 
                 st.markdown("""
                        <div class="centered-container">
                        <h4 class="centered-header">Probability of Travelling Supporters</h4>
                        </div>
                        """, unsafe_allow_html=True)
            
                 numeric_data = pd.to_numeric(simulation_results['TotalTravelDemand'])
                 fig = px.histogram(
                      numeric_data, 
                      histnorm='percent', 
                      nbins=10,
                      #title="Probability of Travelling Supporters"
                      )
                 fig.update_layout(
                      xaxis_title=None,
                      yaxis_title=None,
                      yaxis=dict(range=[0, 100],
                                tickvals=[0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
                                ticktext=['0%', '10%', '20%', '30%', '40%', '50%', '60%', '70%', '80%', '90%', '100%']),
                      showlegend=False,
                      #width=50,   
                      #height=350, 
                      #autosize=True,
                      bargap=0.2)
                 
                 st.plotly_chart(fig, use_container_width=True)

                 is_button_pressed = 0 #reset the button press to 0 to allow reruns


if simulation_mode == "Free":
    freesimcol1, freesimcol2, freesimcol3 = st.columns(3)  
    with freesimcol1:
        st.markdown(f"#### **Simulated Data**")
        df_free_simulation = pd.DataFrame([
            {"Team": "",
             "competition_champion_prob": "",
             "traveling_demand": ""}
             ])
        edited_df = st.data_editor(df_free_simulation, num_rows="dynamic")
        edited_df['Team'] = edited_df['Team'].astype(str)
        edited_df['competition_champion_prob'] = pd.to_numeric(edited_df['competition_champion_prob'])
        edited_df['competition_champion_prob'] = edited_df['competition_champion_prob'] / 100
        edited_df['traveling_demand'] = pd.to_numeric(edited_df['traveling_demand'])


if is_button_pressed == 1 and simulation_mode == "Free":
                
        simulation = []
        matchups = []
        
        np.random.seed(99)
        
        for _ in range(trials):
            finalists = edited_df.sample(n=2, weights='competition_champion_prob')
            finalists = finalists.sort_values(by='Team')
            team1, team2 = finalists['Team'].values
            travel_demand1, travel_demand2 = finalists['traveling_demand'].values
            total_travel_demand = travel_demand1 + travel_demand2
            
            simulation.append({
                 'Team1': team1,
                 'Team2': team2,
                 'TravelDemand1': travel_demand1,
                 'TravelDemand2': travel_demand2,
                 'TotalTravelDemand': total_travel_demand
                 })
            simulation_results = pd.DataFrame(simulation)
            
            matchup = (f"{team1} vs {team2}")
            matchups.append(matchup)
        #calculate the most common matchup
        matchups = pd.DataFrame(matchups, columns = ['match'] )
        matchups_percentages = pd.DataFrame((matchups['match'].value_counts() / trials) * 100).rename(columns={'count': 'percentage'})
        matchups = matchups.drop_duplicates(subset=['match'])
        matchups_percentages = matchups_percentages.merge(matchups[['match']], on='match', how='left')
        top_matchups = matchups_percentages.sort_values(by='percentage', ascending=False).head(5)
        
        with freesimcol2:    
            if is_button_pressed == 1 and simulation_mode == "Free":
                # calculate average travel demand    
                average_demand = round(simulation_results['TotalTravelDemand'].mean())
                st.markdown(f"#### **Estimated Average Traveling Demand:** {average_demand}")
        
        with freesimcol3:    
            if is_button_pressed == 1 and simulation_mode == "Free":
                st.markdown(f"#### **Most Frequent Matchups**")
                top_matchups


                
