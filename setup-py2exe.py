from distutils.core import setup

import py2exe


setup(
    windows=['GradeGUI.py'],
    data_files=[('rule', [r'./rule/new_rule.json',
                          r'./rule/origin_rule.json'])]
)
