#Treating experimental COSMED data from 身体行動計測演習 (resting metabolic rest trial)

import argparse
from pathlib import Path
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


###################################
# arguments: --folder and --verbose
def parse_args():
    parser = argparse.ArgumentParser(description="RMR calculation pipeline")

    parser.add_argument(
        "--folder",
        type=str,
        help="Path to experiment folder"
    )

    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable verbose mode"
    )

    return parser.parse_args()


###################################
# functions:
from pathlib import Path # to manage paths

import numpy as np # for array operations
import pandas as pd # for dataframe
import matplotlib.pyplot as plt # for plots
import matplotlib.dates as mdates # for tick interval settings

from datetime import datetime # for operations on date and time
from datetime import timedelta

# getting the path to cosmed data
def get_path(folder, verbose = False):
    """
    get a path to cosmed data
    folder: a string corresponding to the folder where experiment data are located
    return a Path object to the cosmed csv file
    """
    directory = Path.cwd() / "data" / folder
    print("Directory : ", directory)
    dir_content = []
    for item in directory.iterdir():
        dir_content.append(item)
        if ".csv" in str(item):
            path = item
    print(f"content of {folder}/ :")
    for c in dir_content:
        print("    ", c)
    print("Path : ", path)
    return path

def pick_meta(path):
    """
    open cosmed file and pick-up metadata information (participant weight, date, and starting record time)
    path: a Path object leading to a cosmed file
    return:
        - weight: participant weight (int)
        - day: date of the experiment (str)
        - start: the time the cosmed device starts to record (str)
    """
    
    # need update to shut down Warnings
    data = pd.read_csv(path)
    data = data.set_index(data.iloc[:,0])
    meta = {"weight" : int(data.loc["Weight"].iloc[1]),
            "day" : data.loc["Date"].iloc[1],
            "start" : data.loc["Time"].iloc[1]
           }
   
    print("subject weight is :", meta["weight"])
    print("date of experiment :", meta["day"])
    print("actual starting time :", meta["start"])

    return meta

def format_time(path, meta, verbose = False):
    """
    open cosmed file, format time information, set time as index
    path: a Path object leading to a cosmed file
    meta: a list with pariticipant weight
    return:
        - a panda dataframe with all cosmed parameters and datetime objects in index
    """
    
    # read cosmed data csv one more time and slice to keep time and gas exchange parameter only
    data = pd.read_csv(path, skiprows=26)
    data.columns = data.columns.str.strip() # remove fantom space in column names
    data = data.iloc[1:,:]
    
    # make a datetime object out of the string in "hh:mm:ss" 
    elapsed = [datetime.strptime(str(d), "%H:%M:%S").time() for d in data["hh:mm:ss"]]
    elapsed  = [timedelta(hours=e.hour,minutes=e.minute,seconds=e.second) for e in elapsed]
    
    # uncomment for mega debug
    # print("elapsed time : ", elapsed)

    # make a time column
    # turn experiment day from string to datetime object -> date
    day = datetime.strptime(meta["day"], "%Y/%m/%d").date()
    # turn starting time from string to datetime object -> hour, minute, second
    start = datetime.strptime(str(meta["start"]), "%H:%M:%S").time()
    # combine date and starting time objects into a single object -> year, month, day - hour, minute, second
    start_day_time = datetime.combine(day, start)
    # add elapsed time to starting date and time and print the length (should be same as data length)
    time = [start_day_time + e for e in elapsed]
    
    if verbose:
        print("day (datetime object) : ", day)
        print("start time (datetime object) : ", start)
        print("start day and time (datetime object) : ", start_day_time)
        print("number of datetime objects : ", len(time))
    print("fromating time information as datetime objects: OK")
    # use time as index
    data.index = time
    print("use as index: OK") 
    return data

def get_start_end_rmr(folder, exp_day, verbose = False):
    """
    get the actual starting and ending time of the experiment
    folder: a string corresponding to the folder where experiment data are located
    exp_day: a datetime object corresponding to the experiment day.
    return: a dictionary of 2 datetime objects
    """
    directory = Path.cwd() / "data" / folder
    print("Directory : ", directory)

    dir_content = list(directory.iterdir())
    # Check if any Excel file exists
    has_excel = any(item.suffix == ".xlsx" for item in dir_content)
    if not has_excel:
        print("No Excel file found in directory.")
        start_time = input("Please enter a starting time for the rmr experiment. Format: hh:mm:ss")
        end_time = input("Please enter a starting time for the rmr experiment. Format: hh:mm:ss")
 
    else:
        for item in dir_content:
            if item.suffix == ".xlsx":
                data = pd.read_excel(item, header=None).astype(str)

                has_start = "start-rmr" in str(data.iloc[0, 0])
                has_end = "end-rmr" in str(data.iloc[1, 0])

                if verbose:
                    print("start-rmr :", has_start)
                    print("end-rmr   :", has_end)

                if has_start and has_end:
                    start_time = data.iloc[0, 1]
                    end_time = data.iloc[1, 1]
                    break  # stop once found
        else:
            print("No valid start/end found in Excel files.")
            start_time = input("Please enter a starting time for the rmr experiment. Format: hh:mm:ss")
            end_time = input("Please enter a starting time for the rmr experiment. Format: hh:mm:ss")
 
    if verbose:
        print("content of directory :")
        for c in dir_content:
            print("    ", c)
    
    # combine date and start and end time objects into a single object -> year, month, day - hour, minute, second
    day = datetime.strptime(exp_day, "%Y/%m/%d").date()
    start = datetime.strptime(start_time, "%H:%M:%S").time()
    end = datetime.strptime(end_time, "%H:%M:%S").time()
    start_day_time = datetime.combine(day, start)
    end_day_time = datetime.combine(day, end)
    print("collecting rmr trial start and end time information: OK")  
    return start_day_time, end_day_time
       
def cut_and_slice(data, start, end, gas_only=True):
    """
    select parameters of interest only and trim data according to experiment start and end information
    parameters:
    data: a pandas dataframe (built from original cosmed csv file) with datetime objects in index.
    start: a datetime object containing the experiment start information.
    end: a datetime object containing the experiment end information.
    return: a trimmed pandas data frame with VO2 and VCO2 time series
    """
    select = data.loc[start:end]
    print("data trimming: OK") 
    select = select.rename(columns={"K5_VO2":"VO2","K5_VCO2":"VCO2"})
    if gas_only:
        select = select[["VO2","VCO2"]]  
        print("data slicing ('VO2','VCO2' only): OK")  
    return select

def plot_data(data, folder, tick_interval=60):
    """
    data: a pandas dataframe
    folder: a string pointing to the data location folder (for plot title).
    tick_interval: (int) a number of row corresponding to interval between ticks
    """
 
    fig, ax = plt.subplots(figsize=(12, 8))

    for c in data.columns:
        plt.plot(data.index, data[c], label=c) #x,y,label
        
    plt.xlabel("date and time")
    plt.ylabel(", ".join(data.columns))
    plt.legend()
    plt.title(folder)

    plt.tight_layout()
    
    # tick every {tick_interval} seconds
    ax = plt.gca()
    ax.xaxis.set_major_locator(mdates.SecondLocator(interval=tick_interval))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m/%d-%H:%M:%S'))

    # all ticks
    #plt.xticks(data.index,rotation=90)

    plt.xticks(rotation=90)    

    plt.show(block=False)
    
    return fig, ax 

def smoothing(data, folder, window_length=60, method="equal-weight", plot=False):
    """
    making a dataset with regular interval of 1 second (interpolating) and smoothing the data for sliding windows of a given length
    parameters:
    data: a pandas dataframe with datetime objects index.
    folder: a string pointing to the data location folder (for optional plot title).
    window_length: an int corresponding to the length of teh window in seconds (default: 60).
    method: "equal-weight" (gaussian) or "gaussian". If "equal-weight", equal weight is given to all points in the window.
    If "gaussian", std =15.
    plot: show a plot if plot = True.
    
    """
    smooth = data.resample("1s").mean(numeric_only=True).interpolate()
    
    for c in smooth.columns:
        if method == "gaussian":
            smooth[new_col] = smooth[c].rolling(window = window_length, center = True, min_periods=1, win_type="gaussian").mean(std=15) 
        elif method == "equal-weight":
            smooth[f"smooth_{c}"] = smooth[c].rolling(window = str(window_length)+"s", center = True, min_periods=1, win_type=None).mean() 
    print("signal smoothing: OK")    
    
    if plot:
        plot_data(smooth, folder)

    return smooth

def rmr_calculation(data, folder):
    """
    Calculate metabolic rate (kcal/d) using Weir equation in different period of time, general stats are extracted.
    data: a dataframe containing smooth VO2 and VCO2 signals and having datetime objects for index.
    folder: a string pointing to the data location folder (for optional plot title).
    return:
        a table (pandas data frame) with rmr calculation result.
        The table contains rmr calculation results for:
            - the lowest 2 and 5 minutes
            - the last 2 and 5 minutes
            - the whole rmr trial period
        RMR: mr calculated every second with or without windowing
        view: a plot object plot ready to be saved
    """
    #　calculate metabolic rate (kcal / day) for each data point
    mr = (3.9*(data["smooth_VO2"])/1000 + 1.1*(data["smooth_VCO2"])/1000)*1440 # VO2 and VCO2 in ml/min ->  L/day
    print("metabolic rate date series: OK")    
    # result for the whole rmr trial
    rmr_summary_whole = {"mean_whole":mr.dropna().mean(), "std_whole":mr.dropna().std(),
               "min_whole":mr.dropna().min(), "min_time_whole": mr.dropna().idxmin(),
               "median_whole":mr.dropna().median(),
               "max_whole":mr.dropna().max(), "max_time_whole": mr.dropna().idxmax(),
               "last_whole":mr.iloc[-1]}
    print("metabolic rate statitistics for the whole length: OK")
    # results for window analysis of 2 min
    mr_2min = mr.rolling(window=120, center = True).mean()
    rmr_summary_2min = {"mean_2":mr_2min.dropna().mean(),"std_2":mr_2min.dropna().std(),
               "min_2":mr_2min.dropna().min(),"min_time_2": mr_2min.dropna().idxmin(),
               "median_2":mr_2min.dropna().median(),
               "max_2":mr_2min.dropna().max(),"max_time_2": mr_2min.dropna().idxmax(),
               "last_2":mr_2min.dropna().iloc[-1]}
    print("metabolic rate statitistics for the 2-min window analysis: OK")
    # results for window analysis of 5 min
    mr_5min = mr.rolling(window=300, center = True).mean()
    rmr_summary_5min = {"mean_5":mr_5min.dropna().mean(),"std_":mr_5min.dropna().std(),
               "min_5":mr_5min.dropna().min(),"min_time_5": mr_5min.dropna().idxmin(),
               "median_5":mr_5min.dropna().median(),
               "max_5":mr_5min.dropna().max(),"max_time_5": mr_5min.dropna().idxmax(),
               "last_5":mr_5min.dropna().iloc[-1]}
    print("metabolic rate statitistics for the 5-min window analysis: OK")
    
    RMR = pd.concat([mr.to_frame(name="RMR_no_window"), mr_2min.to_frame(name="RMR_2min_windows"), mr_5min.to_frame(name="RMR_5min_windows")], axis=1)
    results = pd.DataFrame([{**rmr_summary_whole, **rmr_summary_2min, **rmr_summary_5min}])
    results.insert(0, "exp", folder) # first column = exp id
    view , _ = plot_data(RMR, folder)
    print("all rmr results: OK")    
    return results, RMR, view

def save_result(table, image, folder):
    """
    saving routine
    table: a pandas dataframe with results.
    image: a plt figure object.
    folder: data location
    return a csv with rmr results and a png with rmr kinetics
    """
    now = datetime.now().strftime("%Y%m%d-%H%M%S")
    base_dir = Path("results")
    base_dir.mkdir(parents=True, exist_ok=True)  # ensure results/ exists or create it
    
    # save table results
    csv_name = now+"-table.csv"
    csv_path = base_dir / csv_name
    table.to_csv(csv_path, index=False)
    print(f"Results saved in: {csv_path}")
    # save image
    image_name = now+"-"+folder+".png"
    image_dir = base_dir / (now+"-images")
    image_dir.mkdir(parents=True, exist_ok = True) # create subfolder
    
    image_path = image_dir / image_name
    image.savefig(image_path, dpi = 300)
    print(f"Image saved in: {image_path}")


###################################
# main:
def main(folder, verbose=False):
    folder_path = Path(folder)

    if not folder_path.exists():
        folder_path = Path.cwd() / "data" / folder

    if not folder_path.exists():
        raise FileNotFoundError(f"Folder not found: {folder}")

    print(f"Processing folder: {folder}")
    print(f"Verbose: {verbose}")

    path = get_path(folder)
    meta = pick_meta(path)
    data = format_time(path, meta, verbose)
    start, end = get_start_end_rmr(folder, meta["day"], verbose)
    selected = cut_and_slice(data, start, end)
    smooth = smoothing(selected, folder, method="equal-weight", plot=True)
    result, _ , figure = rmr_calculation(smooth, folder)
    save_result(result,figure,folder)


###################################
# entry point
if __name__ == "__main__":
    args = parse_args()

    folder = args.folder if args.folder else input("Enter folder: ")
    verbose = args.verbose

    main(folder, verbose)