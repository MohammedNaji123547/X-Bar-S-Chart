import pandas as pd


def get_random_sample(source_file):
    df = pd.read_excel(source_file, 'Lead Time')
    new_df = df.dropna()

    shortest_column_length = len(new_df)

    result_df = pd.DataFrame()

    for (columnName, columnData) in df.items():
        sample = df[columnName].dropna().sample(n=shortest_column_length)
        result_df[columnName] = sample.reset_index()[columnName]

    return result_df


# def duration_to_seconds(duration):
#     hours, minutes, seconds = map(int, duration.split(':'))
#     total_seconds = hours * 3600 + minutes * 60 + seconds
#     return total_seconds
