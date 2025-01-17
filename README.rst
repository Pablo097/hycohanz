hycohanz
========

hycohanz_ is an Open Source (BSD license) Python wrapper interface to the ANSYS HFSS Windows COM API,
enabling you to control HFSS from Python.
hycohanz simplifies control of HFSS from Python for RF, microwave, and antenna engineers.

.. _hycohanz:  http://mradway.github.io/hycohanz/

DISCONTINUATION NOTICE
----------------------

hycohanz will probably be discontinued soon, as an official Python library for Ansys products already exists, called **PyAnsys**. Its website https://docs.pyansys.com/ contains documentation about it, the projects are open-source and the code for the AEDT-related library (PyAEDT) can be found `here <https://github.com/pyansys/pyaedt>`_. This official library is much more complete, pythonic and supported than hycohanz will ever be, so it is recommended to use PyAEDT instead.

Even so, if you find this library easier to use for your project and you would like any lacking HFSS function or improvement to be implemented, feel free to post it in the `issue tracker <https://github.com/Pablo097/hycohanz/issues>`_ and I will try to implement it as soon as possible.

Minimal Example
---------------

.. sourcecode:: python

    import hycohanz as hfss

    [oAnsoftApp, oDesktop] = hfss.setup_interface()

    input('Press "Enter" to quit HFSS.>')

    hfss.quit_application(oDesktop)

    del oDesktop
    del oAnsoftApp

Dozens more examples_ are included in the examples directory of the source distribution.

.. _examples:  https://github.com/Pablo097/hycohanz/tree/devel/examples

However, for the moment, the majority of these examples are from the original repository, 
so they do not incorporate most of the newer functions. Nevertheless, they are enough to
understand how the library works.

This library now supports calling every function without their COM object first arguments
(i.e. oAnsoftApp, oDesktop, oProject, oEditor and oDesign, correspondingly),
as their current values are handled and stored internally. That is to say that
the previous minimal example can be simplified to:

.. sourcecode:: python

    import hycohanz as hfss

    hfss.setup_interface()

    input('Press "Enter" to quit HFSS.>')

    hfss.quit_application()

    hfss.clean_interface()


Quick Install
-------------

Installation is easy if you already have HFSS_ and Python_. hycohanz uses
the pywin32_ Windows extensions for Python and the quantiphy_ library. All of
these are automatically installed (if not already) when installing hycohanz.

.. _HFSS: http://www.ansys.com/Products/Simulation+Technology/Electromagnetics/Signal+Integrity/ANSYS+HFSS
.. _Python:  http://www.python.org
.. _pywin32:  https://github.com/mhammond/pywin32
.. _quantiphy:  https://quantiphy.readthedocs.io/en/stable/

1. Download the `.zip file`_ from Github.

.. _`.zip file`:  https://github.com/Pablo097/hycohanz/archive/devel.zip

2. Unzip to a convenient location.

3. Open a Windows command shell prompt in that location and run::

    > python -m pip install .

In case you don´t already have Python installed, you can download it from here_.

.. _here: https://www.python.org/downloads/

Alternatively, if you are not interested in having a local copy of the project, the library can be installed 
without downloading it manually, simply executing the following command in any Windows command shell prompt::
    > python -m pip install https://github.com/Pablo097/hycohanz/archive/devel.zip
    
Updating
--------

In order to update your local library installation with newer available versions, you can either re-download 
the project and follow the three-step process from the previous section, or execute the following command in 
a command shell prompt::
    > python -m pip install --force-reinstall --no-deps https://github.com/Pablo097/hycohanz/archive/devel.zip

Problems, Bugs, Questions, and Feature Requests
-----------------------------------------------
These are currently handled via the hycohanz issue tracker https://github.com/Pablo097/hycohanz/issues.

This issue tracker is useful if you

- run into problems with installing or running hycohanz,
- find a bug,
- have a question,
- would like to see a feature implemented.

Of course, you can also email_ me privately.

.. _email:  mailto:pablomr@ic.uma.es

Features
--------
hycohanz provides convenience functions for the following:

- Starting, connecting to, and closing HFSS
- Creating design variables
- Manipulating HFSS expressions
- Creating 3D models using polylines, circles, rectangles, spheres, etc.
- Querying objects and groups of objects
- Object manipulation via unite, subtract, imprint, mirror, move, cut, paste, rotate, scale, sweep, etc.
- Assigning boundary conditions and excitations
- Manipulating projects and designs
- Creating analysis setups and frequency sweeps
- Create reports and export data

In addition, most of the functions support feeding them with numeric values or HFSS expressions (strings with 
equations using HFSS design variables) indistinctly. This is true even for the start and end coordinates of the 
integration lines for the excitation assignment functions (waveport, lumpedport...), which are known to only admit
explicitly numeric values by default.

Examples
--------
Dozens of examples_ are included in the examples directory of the source distribution.

.. _examples:  https://github.com/Pablo097/hycohanz/tree/devel/examples

Warning
-------

hycohanz is pre-alpha software and is in active development.
The hycohanz function interfaces can be expected to change frequently, with little concern for backwards compatibility.
This situation is expected to resolve as the project approaches a more mature state.
However, if today you require a stable, reliable, and correct function library for HFSS, unfortunately this library is probably not for you in its current form.

See Also
--------
scikit-rf_:  An actively-developed library for performing common tasks in RF, providing functionality analogous to that provided by the MATLAB RF Toolbox.  If you're working with RF or microwave you should consider getting it.

PyVISA_:  Enables control of instrumentation via Python.

matplotlib_:  Excellent Python 2-D plotting library.

numpy_:  Fundamental functions for manipulating arrays and matrices and performing linear algebra in Python.

scipy_:  Builds upon numpy_ to enable MATLAB-like functionality in Python.

sympy_:  Implements analogous functionality to the MATLAB Symbolic Toolbox.

.. _scikit-rf:  http://scikit-rf.org/
.. _PyVISA:  http://pyvisa.sourceforge.net/
.. _matplotlib:  http://matplotlib.org/
.. _numpy:  http://www.numpy.org/
.. _scipy:  http://www.scipy.org/
.. _sympy:  http://sympy.org/en/index.html

Download
--------

A zip file of the development branch can be downloaded from
https://github.com/Pablo097/hycohanz/archive/devel.zip

Of course, one can also pull the source tree in the usual way using git.

Documentation
-------------

Several basic examples can be found in the examples directory.

Most wrapper functions are documented with useful docstrings, and in most
cases their interfaces tend to follow the HFSS API fairly closely.

For best use of this library you should familiarize yourself with the
information in the HFSS Scripting Guide, available in the HFSS GUI under
Help->Scripting Contents.  The library is intended to be used in consultation
with this resource.

If the docstrings and examples are not sufficient, you will find that
many functions consist of five or fewer lines of simple (almost trivial)
code that are easily understood.

Frequently Asked Questions
--------------------------

:Q: Why not write scripts using Visual Basic for Applications (VBA) or JavaScript (JS)?
:A: I've found that programming in Python is generally much, much easier and more
    powerful than in either of these languages.  Plus, I've generally found that
    Visual Basic scripts run inside HFSS tend to break without useful error
    messages, or worse, crash HFSS entirely.  hycohanz can also crash HFSS. But
    when it does, the Python interpreter gives you a nice stack trace, allowing
    you to determine what went wrong.

:Q: Why use Windows COM instead of .NET?
:A: As I understand it, the Visual Basic examples in the HFSS Scripting Guide
    use Windows COM, so that's what I use.  If you're using IronPython, then
    accessing .NET resources should be trivial.  However, I don't use IronPython
    since I make extensive use in my daily work of numpy, scipy, matplotlib,
    h5py, etc., and IronPython has had issues integrating with these tools
    in the past.

:Q: Why not metaprogram VBA or JS?  Then I could use this library on Linux.
:A: That was my initial approach, because I wanted cross-platform capability.
    Compared to the Windows COM approach, it's a lot more time-consuming, and
    it has all of the drawbacks of the first question.

:Q: Why did you use Python instead of MATLAB?
:A: I'm a recent convert to Python, so I now use Python in my daily workflow
    whenever it's convenient (that means about 99.9% of the time). Python
    gives you keyword arguments, which helps keep the average length in characters
    of a hycohanz function call to a minimum, while minimizing implementation
    overhead compared to MATLAB.

:Q: Why not skip the HFSS interface entirely and directly emit a .hfss file?  Then
    I could use this library on Linux.
:A: I've also considered this approach.  As you may know, .hfss files are
    quasi-human-readable text files with a file format that could in principle be
    reasonably parsed and emitted.  However, the expected implementation effort
    would have been quite a bit higher than I wanted.  Not to mention that the format is not
    (to my knowledge) static, nor is it publicly specified or documented.  Thus, an
    implementation of this approach would be expected to be fragile, crash HFSS
    frequently, and leave non-useful error messages.

Contributing
------------

Often one finds that this library is missing a wrapper for a particular
function.  Fortunately it's often quite easy to add, usually taking
only a few minutes.  Most of the time it's a quick modification of
an existing function.  Many functions can be implemented in five
lines of code or less.  If you do add a feature to the code, please
consider contributing it back to this project.
