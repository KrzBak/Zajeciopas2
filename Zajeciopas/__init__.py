# -*- coding: utf-8 -*-
def classFactory(iface):
    from .Zajeciopas import TestPlugin
    return TestPlugin(iface)
