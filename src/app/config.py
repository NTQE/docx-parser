def configuration():
    """defines the variables for running the script

    :return: basic configuration variables for the script
    """
    base_path = "C:\\MPSAC\\TEST\\CONVERSION"
    area = "Waste Water Treatment"
    if area:
        path = f"{base_path}\\{area}"
        template_path = f"{path}\\TEMPLATE\\EMPTY_TEMPLATE.docx"
        new_path = f"{path}\\NEW"
    else:
        path = base_path
        template_path = f"{path}\\TEMPLATE\\EMPTY_TEMPLATE.docx"
        new_path = f"{path}\\NEW"
    return path, template_path, new_path
