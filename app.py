import base64
import io
from dash import Dash, html, dcc, Input, Output, callback, dash_table, State
from process import process_addresses, generate_excel
import dash_bootstrap_components as dbc

# Use Bootstrap theme for a professional look
app = Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])
server = app.server
app.title = "Tách địa chỉ"

# Custom CSS for additional styling
custom_styles = {
    'header': {
        'backgroundColor': '#4b6584',
        'color': 'white',
        'padding': '2rem',
        'marginBottom': '2rem',
        'borderRadius': '0 0 10px 10px',
        'boxShadow': '0 2px 5px rgba(0,0,0,0.1)'
    },
    'card': {
        'borderRadius': '8px',
        'boxShadow': '0 4px 8px rgba(0,0,0,0.1)',
        'padding': '20px',
        'margin': '20px 0',
        'backgroundColor': 'white'
    },
    'footer': {
        'textAlign': 'center',
        'padding': '20px',
        'marginTop': '30px',
        'borderTop': '1px solid #eee'
    },
    'button': {
        'backgroundColor': '#4b6584',
        'borderColor': '#4b6584',
        'marginTop': '15px',
        'cursor': 'pointer'
    }
}

app.layout = dbc.Container([
    # Header
    html.Div([
        html.H1("TÁCH ĐỊA CHỈ", className="display-4"),
        html.P("Tải file chứa 1 cột địa chỉ cần tách",
               className="lead")
    ], style=custom_styles['header']),

    # Main content
    dbc.Row([
        dbc.Col([
            # Upload card
            dbc.Card([
                dbc.CardHeader("Step 1: Upload Data", className="h5"),
                dbc.CardBody([
                    dcc.Upload(
                        id='upload-data',
                        children=html.Div([
                            html.I(className="fas fa-cloud-upload-alt me-2"),
                            'Drag and Drop or ',
                            html.A('Select Excel File', className="text-primary")
                        ]),
                        style={
                            'width': '100%',
                            'height': '100px',
                            'lineHeight': '100px',
                            'borderWidth': '2px',
                            'borderStyle': 'dashed',
                            'borderRadius': '10px',
                            'textAlign': 'center',
                            'backgroundColor': '#f8f9fa',
                            'cursor': 'pointer'
                        },
                        multiple=False
                    ),
                    html.Div(id="upload-status", className="mt-3 text-center")
                ])
            ], className="mb-4"),

            # Processing status
            dbc.Card([
                dbc.CardHeader("Step 2: Processing", className="h5"),
                dbc.CardBody([
                    html.Div(id="processing-status"),
                    dbc.Spinner(html.Div(id="loading-output"), color="primary", type="grow"),
                    html.Div([
                        dbc.Button("Download Processed Data",
                                   id="btn-download",
                                   color="primary",
                                   disabled=True,
                                   className="mt-3"),
                        dcc.Download(id="download-excel")
                    ], className="d-grid gap-2 col-6 mx-auto mt-3")
                ])
            ])
        ], width=12)
    ]),

    # Results preview
    dbc.Row([
        dbc.Col([
            html.Div(id="preview-container", className="mt-4")
        ])
    ]),

    # Footer
    html.Footer([
        html.P("© 2025 Address Processing Tool", className="text-muted")
    ], style=custom_styles['footer'])

], fluid=True, className="bg-light min-vh-100 pb-5")


@callback(
    Output('upload-status', 'children'),
    Output('processing-status', 'children'),
    Output('preview-container', 'children'),
    Output('btn-download', 'disabled'),
    Output('loading-output', 'children'),
    Input('upload-data', 'contents'),
    State('upload-data', 'filename')
)
def update_upload_status(contents, filename):
    if not contents:
        return (
            None,
            "Waiting for file upload...",
            None,
            True,
            None
        )

    try:
        # Show upload status
        upload_status = html.Div([
            html.I(className="fas fa-check-circle text-success me-2"),
            f"File uploaded: {filename}"
        ])

        # Decode uploaded file content
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)

        # Process addresses using uploaded file
        result_df = process_addresses(io.BytesIO(decoded))

        # Processing success message
        processing_status = html.Div([
            html.I(className="fas fa-check-circle text-success me-2"),
            f"Processing complete! {len(result_df)} addresses processed.",
            html.Div([
                html.Span(f"Found issues in {result_df['Check'].value_counts().get('Cần kiểm tra', 0)} addresses",
                          className="text-warning" if result_df['Check'].value_counts().get('Cần kiểm tra',
                                                                                            0) > 0 else "")
            ], className="mt-2")
        ])

        # Create preview with styled table
        preview_component = dbc.Card([
            dbc.CardHeader("Data Preview", className="h5"),
            dbc.CardBody([
                dash_table.DataTable(
                    data=result_df.head(10).to_dict('records'),
                    columns=[{'name': i, 'id': i} for i in result_df.columns],
                    page_size=10,
                    style_table={'overflowX': 'auto'},
                    style_header={
                        'backgroundColor': '#4b6584',
                        'color': 'white',
                        'fontWeight': 'bold'
                    },
                    style_cell={
                        'textAlign': 'left',
                        'padding': '8px',
                        'minWidth': '100px',
                    },
                    style_data_conditional=[
                        {
                            'if': {'column_id': 'Check', 'filter_query': '{Check} contains "Cần kiểm tra"'},
                            'backgroundColor': '#ffeaa7',
                            'color': '#d35400'
                        }
                    ]
                ),
                html.P(f"Showing 10 of {len(result_df)} rows", className="text-muted mt-3")
            ])
        ])

        # Store the processed data for download
        global processed_data
        processed_data = generate_excel(result_df)

        return upload_status, processing_status, preview_component, False, None

    except Exception as e:
        error_message = html.Div([
            html.I(className="fas fa-exclamation-triangle text-danger me-2"),
            "Error processing file: " + str(e)
        ])
        return html.Div([
            html.I(className="fas fa-times-circle text-danger me-2"),
            "Upload failed"
        ]), error_message, None, True, None


@callback(
    Output("download-excel", "data"),
    Input("btn-download", "n_clicks"),
    prevent_initial_call=True
)
def download_processed_file(n_clicks):
    if n_clicks:
        return dcc.send_bytes(
            processed_data.getvalue(),
            filename="processed_addresses.xlsx"
        )


# Run the app
if __name__ == "__main__":
    app.run(debug=True)
