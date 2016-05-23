# intersphinx-xlwsf

[Intersphinx](http://www.sphinx-doc.org/en/stable/ext/intersphinx.html) `objects.inv` file providing 
references for [Microsoft Excel worksheet functions](https://support.office.com/en-us/article/Excel-functions-by-category-5F91F4E9-7B42-46D2-9BD1-63F26A86C0EB).
No custom Sphinx domain was created to house these functions; they currently reside in
`:py:function:`.

The recommended dictionary element to add to `intersphinx_mapping` in `conf.py` is:

    'xlwsf': ('https://support.office.com/en-us',
              'https://raw.githubusercontent.com/bskinn/intersphinx-xlwsf/master/objects.inv')
