"""This module provides a command-line interface using click with subcommands.

The main function serves as the entry point.
"""

import base64
import calendar
import io
import pathlib
import tempfile
import typing as t

import click
import googleapiclient.errors  # type: ignore
import googleapiclient.http  # type: ignore
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pdfkit  # type: ignore
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches
from google.oauth2 import service_account
from googleapiclient.discovery import build  # type: ignore

if pathlib.Path(".env").is_file():
    import dotenv

    dotenv.load_dotenv()


def get_credentials_from_env_var(
    env_var: str, scopes: t.List[str]
) -> service_account.Credentials:
    """Decode the base64 encoded service account key from an env var.

    Parameters
    ----------
    env_var
        The env var containing the base64 encoded service account key.
    scopes
        The scopes for which the credentials are required.

    Returns
    -------
    google.oauth2.service_account.Credentials
        The credentials object.
    """
    decoded_key = base64.b64decode(env_var)
    with tempfile.NamedTemporaryFile(delete=False) as temp_file:
        temp_file.write(decoded_key)
        temp_file.flush()
        credentials = service_account.Credentials.from_service_account_file(
            temp_file.name, scopes=scopes
        )
    return credentials


def dataframe_to_word(df, docx_file_path):
    document = Document()

    # Set custom margins (e.g., 0.5 inches for top and bottom)
    sections = document.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)

    # Add a table with borders
    table = document.add_table(rows=1, cols=len(df.columns))
    table.style = "Table Grid"  # Use a built-in style with borders

    # Add header row
    hdr_cells = table.rows[0].cells
    for i, column in enumerate(df.columns):
        hdr_cells[i].text = str(column)

    # Add data rows
    for index, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, value in enumerate(row):
            row_cells[i].text = str(value)

    # Apply borders to each cell (if needed)
    for row in table.rows:
        for cell in row.cells:
            cell._element.get_or_add_tcPr().append(
                parse_xml(r'<w:shd {} w:fill="FFFFFF"/>'.format(nsdecls("w")))
            )
            cell._element.get_or_add_tcPr().append(
                parse_xml(
                    (
                        r'<w:tcBorders %s><w:top w:val="single" w:sz="4"/>'
                        r'<w:left w:val="single" w:sz="4"/>'
                        r'<w:bottom w:val="single" w:sz="4"/>'
                        r'<w:right w:val="single" w:sz="4"/>'
                        r"</w:tcBorders>"
                    )
                    % nsdecls("w")
                )
            )

    document.save(docx_file_path)


@click.group()
def main() -> None:
    """Main entry point for the Click command-line interface."""
    ...


@main.command()
@click.option(
    "--google-workspace-service-account-key",
    envvar="GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY",
    required=True,
    type=str,
    help=(
        "Base 64 encoded string of the Google Workspace service account key"
        " (JSON) file content. Can be passed as an environment variable"
        " `GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY`. You can set the environment"
        " variable by running"
        " `export GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY=$(base64 key.json)`."
    ),
)
@click.option(
    "--input-file-id",
    envvar="INPUT_FILE_ID",
    required=True,
    type=str,
    help=(
        "Google Drive file id of the input file to download."
        " Can be passed as an environment variable"
        " `INPUT_FILE_ID`."
    ),
)
@click.argument("output_file_path", type=str)
def download_input_file(
    google_workspace_service_account_key: t.Optional[str],
    input_file_id: str,
    output_file_path: str,
) -> None:
    """Download the input file `HomeKitaTage.xlsx` from Google Drive.

    \b
    Arguments
    ---------
    OUTPUT_FILE_PATH
        Path where the downloaded file will be saved.
    """  # noqa: D301
    if not google_workspace_service_account_key:
        raise click.UsageError(
            "Google Workspace service account key is required."
        )
    scopes = ["https://www.googleapis.com/auth/drive"]
    credentials = get_credentials_from_env_var(
        google_workspace_service_account_key, scopes
    )
    hkt_file_path = pathlib.Path(output_file_path)
    hkt_file_path.unlink(missing_ok=True)
    file: t.Optional[io.BytesIO] = None
    try:
        service = build("drive", "v3", credentials=credentials)
        request = service.files().get_media(fileId=input_file_id)
        file = io.BytesIO()
        downloader = googleapiclient.http.MediaIoBaseDownload(file, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        hkt_file_path.write_bytes(file.getvalue())
        if not hkt_file_path.is_file():
            print(f"File {hkt_file_path} not found.")
    except googleapiclient.errors.HttpError as error:
        print(f"An error occurred: {error}")


@main.command()
@click.option(
    "--google-workspace-service-account-key",
    envvar="GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY",
    required=True,
    type=str,
    help=(
        "Base 64 encoded string of the Google Workspace service account key"
        " (JSON) file content. Can be passed as an environment variable"
        " `GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY`. You can set the environment"
        " variable by running"
        " `export GOOGLE_WORKSPACE_SERVICE_ACCOUNT_KEY=$(base64 key.json)`."
    ),
)
@click.option(
    "--parent-id",
    envvar="PARENT_ID",
    required=True,
    type=str,
    help=(
        "Google Drive directory id for the parent of the file(s) to be"
        " uploaded. Can be passed as an environment variable"
        " `PARENT_ID`."
    ),
)
@click.argument("files", type=click.Path(exists=True), nargs=-1)
def upload_files(
    google_workspace_service_account_key: t.Optional[str],
    parent_id: str,
    files: tuple[str],
) -> None:
    """Upload (a) file(s) to a directory in the Google Drive.

    \b
    Arguments
    ---------
    FILES
        Path to the file(s) to be uploaded.
    """  # noqa: D301
    scopes = ["https://www.googleapis.com/auth/drive"]
    if not google_workspace_service_account_key:
        raise click.UsageError(
            "Google Workspace service account key is required."
        )
    credentials = get_credentials_from_env_var(
        google_workspace_service_account_key, scopes
    )
    try:
        service = build("drive", "v3", credentials=credentials)
        existing_files = (
            service.files()
            .list(fields="nextPageToken, files(id, name)")
            .execute()
        ).get("files", [])
        for file_to_upload in files:
            file_path = pathlib.Path(file_to_upload)
            if existing_files:
                for existing_file in existing_files:
                    if existing_file["name"] == file_path.stem:
                        service.files().delete(
                            fileId=existing_file["id"]
                        ).execute()
                        existing_files.remove(existing_file)
                        break
            file_metadata = {
                "name": file_path.stem,
                "parents": [parent_id],
            }
            media = googleapiclient.http.MediaFileUpload(
                file_path, chunksize=-1
            )
            file = (
                service.files()
                .create(
                    body=file_metadata,
                    media_body=media,
                    fields="id,name,webViewLink",
                )
                .execute()
            )
            print(f"Uploaded {file.get('name')} to {file.get('webViewLink')}.")
    except googleapiclient.errors.HttpError as error:
        print(f"An error occurred: {error}")


def _read_and_validate_input_file(
    input_file_path: pathlib.Path,
) -> pd.DataFrame:
    columns = (
        "Name",
        "Group",
        "Monday\nmorning",
        "Monday\nafternoon",
        "Tuesday\nmorning",
        "Tuesday\nafternoon",
        "Wednesday\nmorning",
        "Wednesday\nafternoon",
        "Thursday\nmorning",
        "Thursday\nafternoon",
        "Friday\nmorning",
        "Friday\nafternoon",
        "Assigned by us?",
        "Comments",
    )
    df = pd.read_excel(input_file_path, engine="openpyxl")
    if not df.columns.to_list() == list(columns):
        raise ValueError(
            "Columns of the input file are not as expected."
            f" Expected ordered set of columns: {columns}"
        )
    return df


@main.command()
@click.argument(
    "input_file",
    type=click.Path(exists=True, path_type=pathlib.Path, dir_okay=False),
)
@click.argument(
    "output_dir",
    type=click.Path(exists=True, path_type=pathlib.Path, file_okay=False),
)
def create_statistics(
    input_file: pathlib.Path,
    output_dir: pathlib.Path,
) -> None:
    """Create statistics from the input file.

    \b
    Arguments
    ---------
    INPUT_FILE
        Path to the input file.
    OUTPUT_DIR
        Path to the directory where the statistics will be saved.
        If the directory does not exist, it will be created.
        When the directory exists, it will be emptied before saving the
        statistics.
    """  # noqa: D301
    print(f"Creating statistics from {input_file}.")
    output_dir.mkdir(parents=True, exist_ok=True)
    for file in output_dir.iterdir():
        if file.is_file():
            file.unlink()
    df = _read_and_validate_input_file(input_file)

    for group, group_df in df.groupby("Group"):
        labels = []
        home_values = []
        not_home_values = []

        for no, day in enumerate(list(calendar.day_name)[:5], start=1):
            morning_col = f"{day}\nmorning"
            afternoon_col = f"{day}\nafternoon"

            morning_home = group_df[morning_col].sum()
            afternoon_home = group_df[afternoon_col].sum()

            morning_total = len(group_df)
            afternoon_total = len(group_df)

            morning_not_home = morning_total - morning_home
            afternoon_not_home = afternoon_total - afternoon_home

            labels.extend([f"{day} morning", f"{day} afternoon"])
            home_values.extend([morning_home, afternoon_home])
            not_home_values.extend([morning_not_home, afternoon_not_home])

        x = np.arange(len(labels))  # the label locations
        width = 0.35  # the width of the bars

        fig, ax = plt.subplots(figsize=(10, 6))
        rects1 = ax.bar(
            x - width / 2,
            home_values,
            width,
            label="At home",
            color="green",
        )
        rects2 = ax.bar(
            x + width / 2,
            not_home_values,
            width,
            label="In KITA",
            color="red",
        )

        # Add some text for labels, title and custom x-axis tick labels, etc.
        ax.set_ylabel("Number of children")
        ax.set_title("Distribution")
        ax.set_xticks(x)
        ax.set_xticklabels(labels, rotation=45, ha="right")
        ax.legend(loc="center left", bbox_to_anchor=(1, 0.5))

        # Adjust the layout to make room for the legend
        plt.subplots_adjust(right=0.75)

        # Add numbers above the bars
        def autolabel(rects):
            """Attach a text label above each bar in *rects*, displaying its height."""
            for rect in rects:
                height = rect.get_height()
                ax.annotate(
                    "{}".format(int(height)),
                    xy=(rect.get_x() + rect.get_width() / 2, height),
                    xytext=(0, 3),  # 3 points vertical offset
                    textcoords="offset points",
                    ha="center",
                    va="bottom",
                )

        autolabel(rects1)
        autolabel(rects2)

        fig.tight_layout()

        # Save the plot to a BytesIO object
        img_data = io.BytesIO()
        plt.savefig(img_data, format="png")
        plt.close(fig)
        img_data.seek(0)
        # store img_data in a file
        img_file = output_dir / f"{group}_daily_distributions.png"
        with img_file.open("wb") as f:
            f.write(img_data.read())


@main.command()
@click.argument(
    "input_file",
    type=click.Path(exists=True, path_type=pathlib.Path, dir_okay=False),
)
@click.argument(
    "output_dir",
    type=click.Path(exists=True, path_type=pathlib.Path, file_okay=False),
)
def create_daily_overviews(
    input_file: pathlib.Path,
    output_dir: pathlib.Path,
) -> None:
    """Create daily overviews from the input file.

    \b
    Arguments
    ---------
    INPUT_FILE
        Path to the input file.
    OUTPUT_DIR
        Path to the directory where the daily overviews will be saved.
        If the directory does not exist, it will be created.
        When the directory exists, it will be emptied before saving the
        statistics.
    """  # noqa: D301
    print(f"Creating statistics from {input_file}.")
    output_dir.mkdir(parents=True, exist_ok=True)
    for file in output_dir.iterdir():
        if file.is_file():
            file.unlink()
    df = _read_and_validate_input_file(input_file)
    # replace empty cells (Come to KITA) with -1.0
    df = df.fillna(-1.0)
    for group, group_df in df.groupby("Group"):
        for no, day in enumerate(list(calendar.day_name)[:5], start=1):
            for value, meaning in [
                (1.0, "Stay at home"),
                (-1.0, "Come to KITA"),
            ]:
                day_df = group_df[
                    (group_df[f"{day}\nmorning"] == value)
                    | (group_df[f"{day}\nafternoon"] == value)
                ]
                day_df = day_df.replace(value, meaning)
                day_df = day_df.replace(1.0, "")
                day_df = day_df.replace(-1.0, "")
                file_name = meaning.replace(" ", "_").lower()
                file_name += f"_{group}_{no}_{day}.html"
                html_file_path = pathlib.Path(output_dir) / file_name
                df = day_df[
                    [
                        "Name",
                        "Group",
                        f"{day}\nmorning",
                        f"{day}\nafternoon",
                    ]
                ].fillna("")
                df.to_html(html_file_path, index=False)
                pdf_file_path = html_file_path.with_suffix(".pdf")
                options = {
                    "encoding": "UTF-8",
                    "user-style-sheet": "style.css",
                }
                pdfkit.from_file(
                    input=str(html_file_path),
                    output_path=str(pdf_file_path),
                    options=options,
                    verbose=False,
                )
                docx_file_path = html_file_path.with_suffix(".docx")
                dataframe_to_word(df, docx_file_path)
                html_file_path.unlink()


if __name__ == "__main__":
    main()
