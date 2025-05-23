# baseline_subtraction
Oftentimes when analyzing spectra from various instruments there will be unwanted background radiation, fluoresence or noise that needs to be removed. When performing qualitative analyses, it is important to have every data set be anchored, so to speak, to the horizontal axis such that comparisons between peak intensities, for instance, can be made accurately. This is especially true when individual data sets have varying degrees of noise.

As such, most scientists and engineers use **Origin** or proprietary software from the instrument manufacturer; like Horiba's **LabSpec**. Unfortunately, I had access to neither. I fancy it as a fun exercise to implement some of the algorithms and routines employed by the aforementioned programs, so not a bad way to spend one's time. I looked into two main techniques, *Weighted Least Squares* and *Whittaker-Henderson Smoothing*. Both can be applied to great success.


## basics
These two techniques will have cost functions which will be minimized to find a staisfactory baseline. The corresponsing expressions are found below:

$$J(\beta) = \sum_{i=1}^{n} w_i (y_i - \mathbf{x}_i^T \beta)^2$$

## usage
Just plop the modules into Excel's VBE, and run the modules with the **prg** prefix with the settings of your choosing.


