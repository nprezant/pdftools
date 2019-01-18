
def layout_items(layout):
    '''
    generator for widgets in a layout
    '''
    for i in range(0, layout.count()):
        widget = layout.itemAt(i).widget()
        yield(widget)