import pandas as pd
import plotly.express as px

# Mock data
data = {'AMBIENTE': ['PROD', 'HML', 'DEV'] * 10}
df_full_servers = pd.DataFrame(data)

if 'AMBIENTE' in df_full_servers.columns:
    df_env = df_full_servers['AMBIENTE'].value_counts().reset_index()
    df_env.columns = ['Ambiente', 'Qtd']
    df_env = df_env.sort_values('Qtd', ascending=True)
    
    print("Dtypes:")
    print(df_env.dtypes)
    print("Data:")
    print(df_env)

    colors_proto = {'DEV': '#3b82f6', 'HML': '#0ea5e9', 'PROD': '#1e3a8a'}

    try:
        print("Attempting Plot...")
        fig_env = px.bar(df_env, x='Qtd', y='Ambiente', text='Qtd', orientation='h', color='Ambiente',
                         color_discrete_map=colors_proto)
        print("Plot Success")
        fig_env.write_html("debug_plot.html")
    except Exception as e:
        print(f"Plot Failed: {e}")
