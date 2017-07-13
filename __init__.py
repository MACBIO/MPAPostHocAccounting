# -*- coding: utf-8 -*-
"""
/***************************************************************************
 MPAPostHocAccounting
                                 A QGIS plugin
 This plugin checks how your MPAs meet your placement objectives
                             -------------------
        begin                : 2017-03-02
        copyright            : (C) 2017 by Jonah Sullivan
        email                : jonahsullivan79@gmail.com
        git sha              : $Format:%H$
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
 This script initializes the plugin, making it known to QGIS.
"""


# noinspection PyPep8Naming
def classFactory(iface):  # pylint: disable=invalid-name
    """Load MPAPostHocAccounting class from file MPAPostHocAccounting.

    :param iface: A QGIS interface instance.
    :type iface: QgsInterface
    """
    #
    from .MPA_postHocAccounting import MPAPostHocAccounting
    return MPAPostHocAccounting(iface)
