import pandas as pd
from sklearn.preprocessing import MinMaxScaler
 
def load_and_process_data(file_path):
    # Load Excel file
    df = pd.read_excel(".\List of use-cases - Sent to students_Öresunds.xlsx")
    
    # Ensure necessary columns exist
    required_columns = ['Titel', 'Ide-ägare', 'AO', 'Område (tex HR, Marknad, Infra, produktion)',
                        'Beskrivning', 'Objective (vilket problem löses)',
                        'Outcome (Vad hoppas man uppnå?)', 'Generella krav (tex ska finnas åtkomst fr teams, ska generera ett mejl vid xyz)',
                        'Tekniska förutsättningar (vilket system används idag, vilken data skapas eller behövs)', 'Acceptanskriterier (när är lösningen "klar")', 'Affärsvärde (business benefit)',
                        'Lösningsförslag/Verktyg (tex chatbott, ändra arbetssätt, power automate-flöde)', 'Data-typ', 'AI-förmåga',
                        'Uppskattad implementationstid (XS =<10h, Small = 11-40h, Medium = 41-200h, Large = >200h)',
                        'Prio(1-4)']
    
    for col in required_columns:
        if col not in df.columns:
            df[col] = "Missing Information"  # Fill missing columns with placeholder
    
    # Fill missing values with placeholder text
    df.fillna("Missing Information", inplace=True)
    return df
 
def map_estimated_time(value):
    time_mapping = {'XS': 4, 'Small': 3, 'Medium': 2, 'Large': 1, 'Missing Information': 0}
    return time_mapping.get(value, 0)
 
def rank_use_cases(df):
    # Convert priority and estimated implementation time to numerical values
    df['Prio'] = pd.to_numeric(df['Prio'], errors='coerce').fillna(0)
    df['Uppskattad implementationstid'] = df['Uppskattad implementationstid (XS =<10h, Small = 11-40h, Medium = 41-200h, Large = >200h)'].apply(map_estimated_time)
    
    # Normalize scores using Min-Max Scaling
    scaler = MinMaxScaler()
    df[['Prio', 'Uppskattad implementationstid']] = scaler.fit_transform(df[['Prio', 'Uppskattad implementationstid']])
    
    # Compute ranking score (higher priority, faster implementation preferred)
    df['Ranking Score'] = df['Prio'] + df['Uppskattad implementationstid']
    
    # Rank based on ranking score
    df = df.sort_values(by='Ranking Score', ascending=False)
    return df
 
def generate_csv_report(df, output_path):
    df.to_csv(output_path, index=False)
    print(f"CSV report generated: {output_path}")
 
# Example usage
file_path = ".\List of use-cases - Sent to students_Öresunds.xlsx"  # Replace with actual file path

output_csv = "use_case_report.csv"
 
df = load_and_process_data(file_path)
df = rank_use_cases(df)
generate_csv_report(df, output_csv)