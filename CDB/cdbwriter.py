from typing import List, Dict, Union

from mailmerge import MailMerge

document = MailMerge("CDB Template.docx")


def get_merge_fields() -> None:
    """
    Get the merge fields present in the template document.

    :return: None
    """
    print(document.get_merge_fields())


def generate_document(page: Dict[str, Union[str, int]], nights: List[Dict[str, str]],
                      instructors: List[Dict[str, str]]) -> None:
    """
    Generate the document by merging the provided data.

    :param page: A dictionary containing the data for merging into the document.
    :param nights: A list of dictionaries representing night data.
    :param instructors: A list of dictionaries representing instructor data.
    :return: None
    """
    document.merge(**page)
    document.merge_rows("Night", nights)
    document.merge_rows("Instructor", instructors)
    document.write(f"CDB {''.join([word[0].upper() for word in page['Client'].split()])} {page['CourseCode']}.docx")


def write_doc(
        duty_manager: str,
        duty_manager_phone: Union[str, int],
        course_director: str,
        course_code: str,
        course_dates: str,
        course_type: str,
        client: str,
        client_students: int,
        client_teams: int,
        client_contact: str,
        client_contact_phone: Union[str, int],
        client_history: str,
        h5_client_manager: str,
        h5_client_manager_phone: Union[str, int],
        equipment_collected_by: str,
        equipment_returned_by: str,
        course_budget_provided: int,
        pre_course_1: str,
        pre_course_2: str,
        pre_course_3: str,
        nights: List[Dict[str, str]],
        instructors: List[Dict[str, str]],
        notes: Union[str, None]) -> None:
    """
    Write the document by merging the provided data.

    :param duty_manager: The duty manager's name.
    :param duty_manager_phone: The duty manager's phone number.
    :param course_director: The course director's name.
    :param course_code: The course code.
    :param course_dates: The course dates.
    :param course_type: The course type.
    :param client: The client name.
    :param client_students: The number of client students.
    :param client_teams: The number of client teams.
    :param client_contact: The client contact's name.
    :param client_contact_phone: The client contact's phone number.
    :param client_history: The client history.
    :param h5_client_manager: The H5 client manager's name.
    :param h5_client_manager_phone: The H5 client manager's phone number.
    :param equipment_collected_by: Who the equipment is collected by.
    :param equipment_returned_by: Who the equipment is returned by.
    :param course_budget_provided: The course budget provided.
    :param pre_course_1: The pre-course information 1.
    :param pre_course_2: The pre-course information 2.
    :param pre_course_3: The pre-course information 3.
    :param nights: A list of dictionaries representing night data.
    :param instructors: A list of dictionaries representing instructor data.
    :param notes: Additional notes for the document.
    :return: None
    """
    # Input validation
    if client_students < 0:
        raise ValueError("Invalid value for client_students. It should be a non-negative integer.")
    if client_teams < 0:
        raise ValueError("Invalid value for client_teams. It should be a non-negative integer.")

    try:
        page = {
            "DutyManager": duty_manager,
            "DutyManagerPhone": duty_manager_phone,
            "CourseDirector": course_director,
            "CourseCode": course_code,
            "CourseDates": course_dates,
            "CourseType": course_type,
            "Client": client,
            "ClientStudents": client_students,
            "ClientTeams": client_teams,
            "ClientContact": client_contact,
            "ClientContactPhone": client_contact_phone,
            "ClientHistory": client_history,
            "H5ClientManager": h5_client_manager,
            "H5ClientManagerPhone": h5_client_manager_phone,
            "EquipmentCollectedBy": equipment_collected_by,
            "EquipmentReturnedBy": equipment_returned_by,
            "CourseBudgetProvided": course_budget_provided,
            "PreCourse1": pre_course_1,
            "PreCourse2": pre_course_2,
            "PreCourse3": pre_course_3,
            "Notes": notes
        }

        generate_document(page, nights, instructors)

    except FileNotFoundError as e:
        print("File not found:", str(e))
    except PermissionError as e:
        print("Permission error:", str(e))
    except KeyError as e:
        print("Key error:", str(e))
    except Exception as e:
        print("An error occurred during document generation:", str(e))
