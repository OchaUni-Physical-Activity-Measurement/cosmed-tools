# Exemple cases

### <code>test-01/</code>:
excel file with time information.

## run with:<br>
<code>python rmr-calculation.py --folder data-sample/test-01</code><br>
<code>python rmr-calculation.py --folder data-sample/test-01 --verbose</code> (with verbose)<br>
<code>python rmr-calculation.py</code> (without argument, enter: "test-01")

### <code>test-02/</code>:
excel file not provided. Starting and ending times must be input manually.

<code>python rmr-calculation.py --folder data-sample/test-02</code><br>
<code>python rmr-calculation.py --folder data-sample/test-2 --verbose</code> (with verbose)<br>
<code>python rmr-calculation.py</code> (without argument, enter: "test-02")

## recursive execution :
<code>python rmr-calculation.py --folder data-sample --recursive</code><br>
<code>python rmr-calculation.py --folder data-sample --verbose --recursive</code><br>