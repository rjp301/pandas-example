import os  # standard Python library for system level functions
import datetime as dt  # standard Python library for datetime operations


# define utility function
# useful if code is needed more than once or behaviour is generic
# define parameter and return types of str (string) (not necessary, but useful for type hinting)
def slice_text(string: str, prefix: str, suffix: str) -> str:
    # index of prefix
    # makes us of TERNARY STATMENT (inline if-else)
    # string.index() gets the index of string within another string
    # first index is 0 if prefix is empty or not in string
    i1 = string.index(prefix) if prefix and prefix in string else 0

    # index of suffix
    i2 = string.index(suffix) if suffix and suffix in string else None

    # return the portion of the string in between the end of prefix and index of suffix
    return string[i1 + len(prefix) : i2]


# filter function to remove any files that have an undesirable filename
def filter_files(filename):
    # hidden files start with .
    if filename.startswith("."):
        return False
    # excel temporary files start with ~
    if filename.startswith("~"):
        return False
    # return true by default to accept all other filenames
    return True


# define function
# function is given a directory of semi-consistently named files
# returns latest filepath and associated date
#
# the date in each filename is parsed and used to sort
# much more reliable than sorting alphabetically
def latest_file(
    path,  # not default value defined, therefore is required
    date_format="%Y %m %d",  # default value defined https://strftime.org
    prefix="",
    suffix="",
) -> tuple:  # return type of function, in this case a tuple (immutable list)
    # use listdir os function to get alls files in a directory
    files = os.listdir(path)
    # pass our filter function and files list to the python filter method
    files = list(filter(filter_files, files))

    file_date = []  # instantiate new empty list
    for file in files:
        # use os splitext function to get filename without extension
        # function returns tuple so use object destructuring on left side of equals sign
        name, _ = os.path.splitext(file)

        # get the part of filename containing data string using our slice_text function
        date_str = slice_text(name, prefix, suffix)

        # enter try catch block which will allow code to continue running even if error thrown
        try:
            # convert data string to actual datetime object using datetime library
            # use date_format defined as function parameter
            # https://strftime.org is good resource for format templates
            date = dt.datetime.strptime(date_str, date_format)
        except ValueError as e:
            # if there's an error print to console and move onto the next item
            print(f"ERROR: {date_str} does not match the format of {date_format}")
            continue

        # add tuple to file_date array
        file_date.append((file, date))

    # get the latest file using the second item in the tuple (in this case the date)
    latest = max(file_date, key=lambda i: i[1])

    # files are just filenames not entire path
    # add directory path to filename to get entire filepath
    fname = os.path.join(path, latest[0])

    # date is second item in tuple
    date = latest[1]

    # use string interpolation to print nice summary to console
    print(
        f"{len(files)} files filtered - latest file is {latest[0]} from {date:%Y %m %d}"
    )

    # return tuple of results
    return fname, date
