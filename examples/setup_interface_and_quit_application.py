import hycohanz as hfss

# Remember: with the current library version, all the oAnsoftApp, oDesktop,
# oProject, oDesign and oEditor objects can be omitted.

## In this example, the difference is explicitly shown. The code commented next
## is valid but more cumbersome:

# [oAnsoftApp, oDesktop] = hfss.setup_interface()
#
# input('Press "Enter" to quit HFSS.>')
#
# hfss.quit_application(oDesktop)
#
# hfss.clean_interface()


## The code commented above is equivalent to:

hfss.setup_interface()

input('Press "Enter" to quit HFSS.>')

hfss.quit_application()

hfss.clean_interface()


# As can be seen, this way it is clearer and cleaner
