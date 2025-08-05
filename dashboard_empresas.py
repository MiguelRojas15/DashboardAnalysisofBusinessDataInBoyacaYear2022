import dash
from dash import dcc, html, Input, Output, dash_table
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

df = pd.read_excel("PROYECTO_BOYACA_EMPRESAS_LIMPIO.xlsx")
print("Archivo cargado correctamente")
print(f"Datos cargados: {len(df)} filas, {len(df.columns)} columnas")

app = dash.Dash(__name__)
app.title = "📊 Dashboard Empresas Boyacá"

# Paleta de colores corporativa inspirada en la teoría del color
COLORS = {
    'background': '#f8fafc',          # Blanco suave
    'surface': '#ffffff',             # Blanco puro
    'primary': '#2563eb',             # Azul corporativo
    'secondary': '#1e40af',           # Azul más oscuro
    'accent': '#059669',              # Verde empresarial
    'warning': '#d97706',             # Naranja profesional
    'text': '#1f2937',               # Gris oscuro para texto
    'text_secondary': '#6b7280',      # Gris medio
    'border': '#e5e7eb',             # Gris claro para bordes
    'shadow': 'rgba(0, 0, 0, 0.1)'   # Sombra suave
}

app.layout = html.Div([
    # Header principal
    html.Div([
        html.H1("📊 DASHBOARD EMPRESAS BOYACÁ", 
                style={
                    'textAlign': 'center',
                    'color': COLORS['primary'],
                    'fontSize': '42px',
                    'fontWeight': '700',
                    'margin': '20px 0 10px 0',
                    'fontFamily': 'Inter, -apple-system, BlinkMacSystemFont, sans-serif',
                    'letterSpacing': '-0.5px'
                }),
        html.P("Análisis Estratégico de Empresas - Departamento de Boyacá",
               style={
                   'textAlign': 'center',
                   'color': COLORS['text_secondary'],
                   'fontSize': '18px',
                   'marginBottom': '30px',
                   'fontWeight': '400',
                   'fontFamily': 'Inter, sans-serif'
               })
    ], style={
        'background': f'linear-gradient(135deg, {COLORS["surface"]} 0%, #f1f5f9 100%)',
        'padding': '40px 30px',
        'borderRadius': '16px',
        'margin': '20px',
        'boxShadow': f'0 4px 20px {COLORS["shadow"]}',
        'border': f'1px solid {COLORS["border"]}'
    }),

    # Sección de filtros - TODOS EN UNA LÍNEA
    html.Div([
        # Filtro por Sector
        html.Div([
            html.Label("Filtrar por Sector:", 
                      style={
                          'color': COLORS['text'], 
                          'fontWeight': '600', 
                          'fontSize': '16px',
                          'marginBottom': '8px',
                          'display': 'block',
                          'fontFamily': 'Inter, sans-serif'
                      }),
            dcc.Dropdown(
                id='sector-dropdown',
                options=[{"label": f"🏢 {s}", "value": s} for s in sorted(df["SectorProductivo"].dropna().unique())],
                value=None,
                placeholder="Selecciona un sector",
                clearable=True,
                style={
                    'fontFamily': 'Inter, sans-serif'
                },
                className='custom-dropdown'
            )
        ], style={
            'width': '28%', 
            'display': 'inline-block', 
            'marginRight': '2%',
            'padding': '20px',
            'backgroundColor': COLORS['surface'],
            'borderRadius': '12px',
            'border': f'1px solid {COLORS["border"]}',
            'boxShadow': f'0 2px 8px {COLORS["shadow"]}',
            'verticalAlign': 'top'
        }),

        # Filtro por Municipio
        html.Div([
            html.Label("Filtrar por Municipio:", 
                      style={
                          'color': COLORS['text'], 
                          'fontWeight': '600', 
                          'fontSize': '16px',
                          'marginBottom': '8px',
                          'display': 'block',
                          'fontFamily': 'Inter, sans-serif'
                      }),
            dcc.Dropdown(
                id='municipio-dropdown',
                options=[{"label": f"📍 {m}", "value": m} for m in sorted(df["Municipio"].dropna().unique())],
                value=None,
                placeholder="Selecciona municipios",
                clearable=True,
                multi=True,
                style={
                    'fontFamily': 'Inter, sans-serif'
                },
                className='custom-dropdown'
            )
        ], style={
            'width': '28%', 
            'display': 'inline-block', 
            'marginRight': '2%',
            'padding': '20px',
            'backgroundColor': COLORS['surface'],
            'borderRadius': '12px',
            'border': f'1px solid {COLORS["border"]}',
            'boxShadow': f'0 2px 8px {COLORS["shadow"]}',
            'verticalAlign': 'top'
        }),

        # Filtro por Ventas
        html.Div([
            html.Label("Rango de Ventas (Millones):", 
                      style={
                          'color': COLORS['text'], 
                          'fontWeight': '600', 
                          'fontSize': '16px',
                          'marginBottom': '15px',
                          'display': 'block',
                          'fontFamily': 'Inter, sans-serif'
                      }),
            dcc.RangeSlider(
                id='ventas-slider',
                min=int(df['Ventas mensuales (Millones)'].min()) if not df['Ventas mensuales (Millones)'].isna().all() else 0,
                max=int(df['Ventas mensuales (Millones)'].max()) if not df['Ventas mensuales (Millones)'].isna().all() else 100,
                step=1,
                marks={
                    int(df['Ventas mensuales (Millones)'].min()) if not df['Ventas mensuales (Millones)'].isna().all() else 0: {
                        'label': f'${int(df["Ventas mensuales (Millones)"].min()) if not df["Ventas mensuales (Millones)"].isna().all() else 0}M',
                        'style': {'color': COLORS['accent'], 'fontWeight': '600', 'fontSize': '12px'}
                    },
                    int(df['Ventas mensuales (Millones)'].max()) if not df['Ventas mensuales (Millones)'].isna().all() else 100: {
                        'label': f'${int(df["Ventas mensuales (Millones)"].max()) if not df["Ventas mensuales (Millones)"].isna().all() else 100}M',
                        'style': {'color': COLORS['accent'], 'fontWeight': '600', 'fontSize': '12px'}
                    }
                },
                value=[
                    int(df['Ventas mensuales (Millones)'].min()) if not df['Ventas mensuales (Millones)'].isna().all() else 0,
                    int(df['Ventas mensuales (Millones)'].max()) if not df['Ventas mensuales (Millones)'].isna().all() else 100
                ],
                tooltip={"placement": "bottom", "always_visible": False},
                className='custom-slider'
            )
        ], style={
            'width': '29%', 
            'display': 'inline-block',
            'padding': '20px',
            'backgroundColor': COLORS['surface'],
            'borderRadius': '12px',
            'border': f'1px solid {COLORS["border"]}',
            'boxShadow': f'0 2px 8px {COLORS["shadow"]}',
            'verticalAlign': 'top'
        })
    ], style={
        'backgroundColor': COLORS['background'],
        'padding': '25px',
        'borderRadius': '16px',
        'margin': '20px',
        'boxShadow': f'0 4px 16px {COLORS["shadow"]}',
        'border': f'1px solid {COLORS["border"]}'
    }),

    # KPIs
    html.Div(id='kpis-container', style={'margin': '20px'}),

    # Gráficos superiores
    html.Div([
        html.Div([
            dcc.Graph(id='histograma', style={'height': '320px'})
        ], style={
            'width': '48%',
            'display': 'inline-block',
            'padding': '8px',
            'verticalAlign': 'top'
        }),
        
        html.Div([
            dcc.Graph(id='piechart', style={'height': '320px'})
        ], style={
            'width': '48%',
            'display': 'inline-block',
            'padding': '8px',
            'verticalAlign': 'top'
        })
    ], style={'margin': '10px 0'}),

    html.Div([
        html.Div([
            dcc.Graph(id='boxplot', style={'height': '320px'})
        ], style={
            'width': '48%',
            'display': 'inline-block',
            'padding': '8px',
            'verticalAlign': 'top'
        }),
        
        html.Div([
            dcc.Graph(id='barchart', style={'height': '320px'})
        ], style={
            'width': '48%',
            'display': 'inline-block',
            'padding': '8px',
            'verticalAlign': 'top'
        })
    ], style={'margin': '10px 0'}),

    # Scatter plot
    html.Div([
        dcc.Graph(id='scatterplot', style={'height': '550px'})
    ], style={'padding': '20px'}),

    # Tabla de datos
    html.Div([
        html.H3("📋 TABLA DE DATOS FILTRADOS", 
                style={
                    'color': COLORS['text'], 
                    'textAlign': 'center', 
                    'marginBottom': '25px',
                    'fontSize': '24px',
                    'fontWeight': '600',
                    'fontFamily': 'Inter, sans-serif'
                }),
        html.Div(id='data-table')
    ], style={
        'backgroundColor': COLORS['surface'],
        'padding': '30px',
        'borderRadius': '16px',
        'margin': '20px',
        'boxShadow': f'0 4px 20px {COLORS["shadow"]}',
        'border': f'1px solid {COLORS["border"]}'
    })

], style={
    'backgroundColor': COLORS['background'],
    'minHeight': '100vh',
    'fontFamily': 'Inter, -apple-system, BlinkMacSystemFont, sans-serif'
})

# CSS personalizado para dropdowns y sliders
app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
        <style>
            .custom-dropdown .Select-control {
                background-color: #ffffff !important;
                border: 2px solid #e5e7eb !important;
                border-radius: 8px !important;
                transition: all 0.2s ease !important;
            }
            .custom-dropdown .Select-control:hover {
                border-color: #2563eb !important;
            }
            .custom-dropdown .Select-control--is-focused {
                border-color: #2563eb !important;
                box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.1) !important;
            }
            .custom-dropdown .Select-menu {
                background-color: #ffffff !important;
                border: 1px solid #e5e7eb !important;
                border-radius: 8px !important;
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1) !important;
            }
            .custom-dropdown .Select-option:hover {
                background-color: #f3f4f6 !important;
            }
            .custom-slider .rc-slider-track {
                background: linear-gradient(90deg, #2563eb, #059669) !important;
                height: 6px !important;
            }
            .custom-slider .rc-slider-rail {
                background-color: #e5e7eb !important;
                height: 6px !important;
            }
            .custom-slider .rc-slider-handle {
                border: 3px solid #2563eb !important;
                background: #ffffff !important;
                box-shadow: 0 2px 8px rgba(37, 99, 235, 0.3) !important;
                width: 20px !important;
                height: 20px !important;
                margin-top: -7px !important;
            }
            .custom-slider .rc-slider-handle:hover {
                box-shadow: 0 4px 12px rgba(37, 99, 235, 0.4) !important;
            }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
'''

@app.callback(
    Output('histograma', 'figure'),
    Output('boxplot', 'figure'),
    Output('piechart', 'figure'),
    Output('barchart', 'figure'),
    Output('scatterplot', 'figure'),
    Output('kpis-container', 'children'),
    Output('data-table', 'children'),
    Input('sector-dropdown', 'value'),
    Input('municipio-dropdown', 'value'),
    Input('ventas-slider', 'value')
)
def update_dashboard(sector, municipios, ventas_range):
    filtered_df = df.copy()
    
    if sector:
        filtered_df = filtered_df[filtered_df["SectorProductivo"] == sector]
    
    if municipios:
        filtered_df = filtered_df[filtered_df["Municipio"].isin(municipios)]
    
    if ventas_range:
        filtered_df = filtered_df[
            (filtered_df['Ventas mensuales (Millones)'] >= ventas_range[0]) &
            (filtered_df['Ventas mensuales (Millones)'] <= ventas_range[1])
        ]

    # Cálculo de KPIs
    total_empresas = len(filtered_df)
    promedio_ventas = filtered_df['Ventas mensuales (Millones)'].mean() if not filtered_df['Ventas mensuales (Millones)'].isna().all() else 0
    total_empleados = filtered_df['Numero empleados'].sum() if not filtered_df['Numero empleados'].isna().all() else 0
    sectores_activos = filtered_df['SectorProductivo'].nunique()

    # KPIs con diseño corporativo
    kpis = html.Div([
        html.Div([
            html.Div([
                html.I(className="fas fa-building", style={'fontSize': '24px', 'color': COLORS['primary'], 'marginBottom': '10px'}),
                html.H2(f"{total_empresas:,}", 
                       style={
                           'color': COLORS['primary'], 
                           'fontSize': '36px', 
                           'margin': '0',
                           'fontWeight': '700',
                           'fontFamily': 'Inter, sans-serif'
                       }),
                html.P("Empresas Registradas", 
                      style={
                          'color': COLORS['text_secondary'], 
                          'fontSize': '14px', 
                          'margin': '5px 0 0 0',
                          'fontWeight': '500',
                          'fontFamily': 'Inter, sans-serif'
                      })
            ], style={'textAlign': 'center'})
        ], style={
            'backgroundColor': COLORS['surface'], 
            'padding': '30px 20px', 
            'borderRadius': '12px', 
            'width': '18%', 
            'display': 'inline-block', 
            'margin': '1%',
            'boxShadow': f'0 4px 12px {COLORS["shadow"]}',
            'border': f'2px solid {COLORS["primary"]}',
            'borderTop': f'4px solid {COLORS["primary"]}',
            'transition': 'transform 0.2s ease'
        }),
        
        html.Div([
            html.Div([
                html.I(className="fas fa-dollar-sign", style={'fontSize': '24px', 'color': COLORS['accent'], 'marginBottom': '10px'}),
                html.H2(f"${promedio_ventas:.1f}M", 
                       style={
                           'color': COLORS['accent'], 
                           'fontSize': '36px', 
                           'margin': '0',
                           'fontWeight': '700',
                           'fontFamily': 'Inter, sans-serif'
                       }),
                html.P("Ventas Promedio", 
                      style={
                          'color': COLORS['text_secondary'], 
                          'fontSize': '14px', 
                          'margin': '5px 0 0 0',
                          'fontWeight': '500',
                          'fontFamily': 'Inter, sans-serif'
                      })
            ], style={'textAlign': 'center'})
        ], style={
            'backgroundColor': COLORS['surface'], 
            'padding': '30px 20px', 
            'borderRadius': '12px', 
            'width': '18%', 
            'display': 'inline-block', 
            'margin': '1%',
            'boxShadow': f'0 4px 12px {COLORS["shadow"]}',
            'border': f'2px solid {COLORS["accent"]}',
            'borderTop': f'4px solid {COLORS["accent"]}'
        }),
        
        html.Div([
            html.Div([
                html.I(className="fas fa-users", style={'fontSize': '24px', 'color': COLORS['secondary'], 'marginBottom': '10px'}),
                html.H2(f"{total_empleados:,}", 
                       style={
                           'color': COLORS['secondary'], 
                           'fontSize': '36px', 
                           'margin': '0',
                           'fontWeight': '700',
                           'fontFamily': 'Inter, sans-serif'
                       }),
                html.P("Total Empleados", 
                      style={
                          'color': COLORS['text_secondary'], 
                          'fontSize': '14px', 
                          'margin': '5px 0 0 0',
                          'fontWeight': '500',
                          'fontFamily': 'Inter, sans-serif'
                      })
            ], style={'textAlign': 'center'})
        ], style={
            'backgroundColor': COLORS['surface'], 
            'padding': '30px 20px', 
            'borderRadius': '12px', 
            'width': '18%', 
            'display': 'inline-block', 
            'margin': '1%',
            'boxShadow': f'0 4px 12px {COLORS["shadow"]}',
            'border': f'2px solid {COLORS["secondary"]}',
            'borderTop': f'4px solid {COLORS["secondary"]}'
        }),
        
        html.Div([
            html.Div([
                html.I(className="fas fa-chart-pie", style={'fontSize': '24px', 'color': COLORS['warning'], 'marginBottom': '10px'}),
                html.H2(f"{sectores_activos}", 
                       style={
                           'color': COLORS['warning'], 
                           'fontSize': '36px', 
                           'margin': '0',
                           'fontWeight': '700',
                           'fontFamily': 'Inter, sans-serif'
                       }),
                html.P("Sectores Activos", 
                      style={
                          'color': COLORS['text_secondary'], 
                          'fontSize': '14px', 
                          'margin': '5px 0 0 0',
                          'fontWeight': '500',
                          'fontFamily': 'Inter, sans-serif'
                      })
            ], style={'textAlign': 'center'})
        ], style={
            'backgroundColor': COLORS['surface'], 
            'padding': '30px 20px', 
            'borderRadius': '12px', 
            'width': '18%', 
            'display': 'inline-block', 
            'margin': '1%',
            'boxShadow': f'0 4px 12px {COLORS["shadow"]}',
            'border': f'2px solid {COLORS["warning"]}',
            'borderTop': f'4px solid {COLORS["warning"]}'
        })
    ], style={
        'display': 'flex',
        'justifyContent': 'center',
        'gap': '20px',
        'flexWrap': 'wrap'
    })

    # Template limpio para gráficos
    template = "plotly_white"
    
    # Histograma
    fig_hist = px.histogram(
        filtered_df, x="Ventas mensuales (Millones)",
        nbins=25, title="📊 Distribución de Ventas Mensuales",
        color_discrete_sequence=[COLORS['primary']],
        template=template
    )
    fig_hist.update_layout(
        plot_bgcolor=COLORS['surface'],
        paper_bgcolor=COLORS['surface'],
        font_color=COLORS['text'],
        title_x=0.5,
        showlegend=False,
        title_font_size=18,
        title_font_color=COLORS['text'],
        title_font_family='Inter',
        font_family='Inter'
    )
    fig_hist.update_traces(opacity=0.8, marker_line_width=1, marker_line_color='white')

    # Boxplot
    fig_box = px.box(
        filtered_df, x="SectorProductivo", y="Ventas mensuales (Millones)",
        title="📈 Análisis de Ventas por Sector Productivo",
        color="SectorProductivo",
        template=template,
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    fig_box.update_layout(
        plot_bgcolor=COLORS['surface'],
        paper_bgcolor=COLORS['surface'],
        font_color=COLORS['text'],
        title_x=0.5,
        showlegend=False,
        title_font_size=18,
        title_font_color=COLORS['text'],
        title_font_family='Inter',
        font_family='Inter'
    )
    fig_box.update_xaxes(tickangle=45)

    # Pie chart
    genero_counts = filtered_df['Genero responsable'].value_counts()
    fig_pie = px.pie(
        names=genero_counts.index, values=genero_counts.values,
        title="👥 Distribución por Género del Responsable",
        color_discrete_sequence=[COLORS['primary'], COLORS['accent'], COLORS['warning'], COLORS['secondary']],
        template=template
    )
    fig_pie.update_traces(
        textposition='inside', 
        textinfo='percent+label', 
        textfont_size=14, 
        marker=dict(line=dict(color='white', width=2)),
        textfont_family='Inter'
    )
    fig_pie.update_layout(
        plot_bgcolor=COLORS['surface'],
        paper_bgcolor=COLORS['surface'],
        font_color=COLORS['text'],
        title_x=0.5,
        title_font_size=18,
        title_font_color=COLORS['text'],
        title_font_family='Inter',
        font_family='Inter'
    )

    # Bar chart
    top_municipios = filtered_df['Municipio'].value_counts().nlargest(10)
    fig_bar = px.bar(
        x=top_municipios.values, y=top_municipios.index,
        orientation='h', title="🏛️ Top 10 Municipios por Número de Empresas",
        color=top_municipios.values,
        color_continuous_scale='Blues',
        template=template
    )
    fig_bar.update_layout(
        plot_bgcolor=COLORS['surface'],
        paper_bgcolor=COLORS['surface'],
        font_color=COLORS['text'],
        title_x=0.5,
        showlegend=False,
        coloraxis_showscale=False,
        title_font_size=18,
        title_font_color=COLORS['text'],
        title_font_family='Inter',
        font_family='Inter'
    )

    # Scatter plot
    fig_scatter = px.scatter(
        filtered_df, x="Numero empleados", y="Ventas mensuales (Millones)",
        color="SectorProductivo", size="Ventas mensuales (Millones)",
        hover_name="NombreEmpresa",
        title="🔍 Relación Empleados vs Ventas por Sector",
        size_max=25,
        template=template,
        opacity=0.7,
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    fig_scatter.update_layout(
        plot_bgcolor=COLORS['surface'],
        paper_bgcolor=COLORS['surface'],
        font_color=COLORS['text'],
        title_x=0.5,
        title_font_size=18,
        title_font_color=COLORS['text'],
        title_font_family='Inter',
        font_family='Inter'
    )

    # Tabla
    table = dash_table.DataTable(
        data=filtered_df.head(20).to_dict('records'),
        columns=[{"name": i, "id": i} for i in filtered_df.columns[:8]],
        style_cell={
            'backgroundColor': COLORS['surface'],
            'color': COLORS['text'],
            'border': f'1px solid {COLORS["border"]}',
            'textAlign': 'left',
            'fontFamily': 'Inter, sans-serif',
            'fontSize': '14px',
            'padding': '12px'
        },
        style_header={
            'backgroundColor': COLORS['primary'],
            'color': 'white',
            'fontWeight': '600',
            'fontSize': '14px',
            'textAlign': 'center',
            'fontFamily': 'Inter, sans-serif'
        },
        style_data_conditional=[
            {
                'if': {'row_index': 'odd'},
                'backgroundColor': '#f8fafc'
            }
        ],
        page_size=10,
        sort_action="native",
        filter_action="native",
        style_table={'overflowX': 'auto', 'borderRadius': '8px'}
    )

    return fig_hist, fig_box, fig_pie, fig_bar, fig_scatter, kpis, table

if __name__ == '__main__':
    print("🚀 Iniciando Dashboard Corporativo...")
    print("🌐 Abre tu navegador en: http://127.0.0.1:8050")
    print("⏹️ Para detener: Ctrl+C")
    app.run(debug=True, host='127.0.0.1', port=8050)