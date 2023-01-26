from src.app import setup, config
import pandas as pd


def main():
    # gathering basic variables for script environment
    path, _, _a = config.configuration()

    print(f"Base Path: {path}")

    # run a function that gathers files in the 'path' directory
    file_path_list = setup.find_files(path)

    # run a function that takes those paths and gathers data from each document
    file_data_list = setup.gather_from_files(file_path_list)

    # create a custom list of lists to convert to csv's for review
    csv_list = []
    lockout_list = []
    for data in file_data_list:
        csv_entry = []
        csv_entry.append(data.file_name)
        csv_entry.append(data.dept)
        csv_entry.append(data.doc_no)
        csv_entry.append(data.equip_id)
        csv_entry.append(data.equip_desc)
        csv_entry.append(data.date_eff)
        csv_entry.append(data.prep_by)
        csv_entry.append(data.prep_by_title)
        csv_entry.append(data.app_by)
        csv_entry.append(data.app_by_title)
        csv_entry.append(data.extra)
        csv_entry.append(data.special_precautions)
        csv_list.append(csv_entry)
        lockout_list.append([data.file_name, "", "", "", "", "", "", ""])
        for lr in data.lockout:
            lockout_list.append(["", lr.num, lr.point, lr.tag_no, lr.energy_src, lr.isolating_means, lr.context, lr.context_text])

    df = pd.DataFrame(csv_list, columns=["file", "dept", "doc_no", "equip_id", "equip_desc", "date_eff", "prep_by", "prep_by_title", "app_by", "app_by_title", "extra", "precautions"])
    df.to_csv(f"{path}\\data.csv", index=False)

    df2 = pd.DataFrame(lockout_list, columns=["file_name", "num", "point", "tag_no", "energy_src", "isolating_means", "context", "context_text"])
    df2.to_csv(f"{path}\\data_lockout.csv", index=False)

    # run a function that inserts the gathered data into an empty template and saves the document
    setup.insert_into_files(file_data_list)


if __name__ == '__main__':
    main()
