#Treating experimental COSMED data from 身体行動計測演習 (resting metabolic rest trial)

import argparse
from pathlib import Path
from datetime import datetime, timedelta
import numpy as np # for array operations
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


###################################
# arguments: --folder and --verbose
def parse_args():
	parser = argparse.ArgumentParser(description="RMR calculation pipeline")

	parser.add_argument(
		"-f", "--folder",
		type=str,
		help="Path to experiment folder"
	)

	parser.add_argument(
		"-v", "--verbose",
		action="store_true",
		help="Enable verbose mode"
	)
	# TODO
	#parser.add_argument(
#		"-m", "--mets",
#		action="store_true",
#		help="Perform METS calculation"
#	)
	parser.add_argument(
		"-r", "--recursive",
		action="store_true",
		help="Process all folders at path")
	return parser.parse_args()


###################################
# functions:

# getting the path to cosmed data
def get_cosmed_filepath(folder, verbose = False):
	"""
	get a path to cosmed data
	folder: a string corresponding to the folder where experiment data are located
	return a Path object to the cosmed csv file
	"""
	directory = Path(folder)
	directory_name = directory.name
	filepath = None
	if verbose:
		print("Directory : ", directory)
	dir_content = []
	for item in directory.iterdir():
		dir_content.append(item)
		if ".csv" in str(item):
			filepath = item
	if verbose:
		print(f"content of {directory_name}/ :")
		for c in dir_content:
			print("    ", c)
		print("Path : ", filepath)
	return filepath

def pick_meta(path, verbose = False):
	"""
	open cosmed file and pick-up metadata information (participant weight, date, and starting record time)
	path: a Path object leading to a cosmed file
	return:
		- weight: participant weight (int)
		- day: date of the experiment (str)
		- start: the time the cosmed device starts to record (str)
	"""
	
	# need update to shut down Warnings
	data = pd.read_csv(path, skiprows=5)
	data = data.set_index(data.iloc[:,0])
	data_meta = data.loc[:"Notes"]
	meta ={str(v).lower(): data.loc[v].iloc[1] for v in data_meta.index}
	#meta = {"weight" : int(data.loc["Weight"].iloc[1]),
	#        "day" : data.loc["Date"].iloc[1],
	#        "start" : data.loc["Time"].iloc[1]
	#       }
	meta["day"] = meta["date"]
	meta["start"] = meta["time"]
	if verbose:
		print(meta)
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
	directory = Path(folder)
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

def plot_data(data, folder, tick_interval=60, verbose=False):
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
	plt.title(Path(folder).name)
	
	# tick every {tick_interval} seconds
	ax = plt.gca()
	ax.xaxis.set_major_locator(mdates.SecondLocator(interval=tick_interval))
	ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y/%m/%d-%H:%M:%S'))

	# all ticks
	#plt.xticks(data.index,rotation=90)

	plt.xticks(rotation=90)    
	plt.tight_layout()

	if verbose:
		plt.show(block=True)
	
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
		plot_data(smooth, folder,verbose=True)

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
	results.insert(0, "exp", Path(folder).name) # first column = exp id
	view , _ = plot_data(RMR, folder,verbose=False) # don't show
	print("all rmr results: OK")    
	return results, RMR, view

def save_image(image, folder, now="test"):
	"""
	saving routine
	image: a plt figure object.
	folder: data location
	"""
	# save image
	
	base_dir = Path("results")
	base_dir.mkdir(parents=True, exist_ok=True)  # ensure results/ exists or create it
	image_name = now+"-"+Path(folder).name+".png"
	image_dir = base_dir / (now+"-images")
	image_dir.mkdir(parents=True, exist_ok = True) # create subfolder
	
	image_path = image_dir / image_name
	image.savefig(image_path, dpi = 300)
	print(f"Image saved in: {image_path}")

def save_result(table, folder,now="test"):
	"""
	saving routine
	table: a pandas dataframe with results.
	folder: data location
	"""
	base_dir = Path("results")
	base_dir.mkdir(parents=True, exist_ok=True)  # ensure results/ exists or create it
	
	# save table results
	csv_name = now+"-table.csv"
	csv_path = base_dir / csv_name
	table.to_csv(csv_path, index=False)
	print(f"Results saved in: {csv_path}")

###################################
# main:
def main(folder, recursive=False, verbose=False):
	now = datetime.now().strftime("%Y%m%d-%H%M%S")
	folder_path = Path(folder)

	if not folder_path.exists():
		folder_path = Path.cwd() / "data" / folder

	if not folder_path.exists():
		raise FileNotFoundError(f"Folder not found: {folder}")

	dirs = []
	if recursive:
		for d in folder_path.iterdir():
			if d.is_dir():
				dirs.append(d)
	else:
		dirs = [folder_path]

	results = []
	print(f"Verbose: {verbose}")
	for f in dirs:
		print(f"Processing folder: {f}")
		path = get_cosmed_filepath(f, verbose)
		meta = pick_meta(path, verbose)
		data = format_time(path, meta, verbose)
		start, end = get_start_end_rmr(f, meta["day"], verbose)
		selected = cut_and_slice(data, start, end)
		smooth = smoothing(selected, f, method="equal-weight", plot=True)
		result, _ , figure = rmr_calculation(smooth, f)
		results.append(result)
		save_image(figure,f,now=now)
		plt.close(figure)
	result = pd.concat(results)
	save_result(result,folder,now=now)


###################################
# entry point
if __name__ == "__main__":
	args = parse_args()

	folder = args.folder if args.folder else input("Enter folder: ")
	verbose = args.verbose
	rec = args.recursive

	main(folder, rec, verbose)