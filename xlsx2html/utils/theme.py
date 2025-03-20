from colorsys import rgb_to_hls, hls_to_rgb
#https://bitbucket.org/openpyxl/openpyxl/issues/987/add-utility-functions-for-colors-to-help

RGBMAX = 0xff  # Corresponds to 255
HLSMAX = 240  # MS excel's tint function expects that HLS is base 240. see:
# https://social.msdn.microsoft.com/Forums/en-US/e9d8c136-6d62-4098-9b1b-dac786149f43/excel-color-tint-algorithm-incorrect?forum=os_binaryfile#d3c2ac95-52e0-476b-86f1-e2a697f24969

def rgb_to_ms_hls(red, green=None, blue=None):
    """Converts rgb values in range (0,1) or a hex string of the form '[#aa]rrggbb' to HLSMAX based HLS, (alpha values are ignored)"""
    try:
        if green is None:
            if isinstance(red, str):
                # Make sure the hex string is valid
                if not all(c in '0123456789abcdefABCDEF' for c in red.replace('#', '')):
                    return (0, HLSMAX, 0)  # Return white instead of gray
                    
                if len(red) > 6:
                    red = red[-6:]  # Ignore preceding '#' and alpha values
                try:
                    blue = int(red[4:], 16) / RGBMAX
                    green = int(red[2:4], 16) / RGBMAX
                    red = int(red[0:2], 16) / RGBMAX
                except (ValueError, IndexError):
                    # If any part fails, return white
                    return (0, HLSMAX, 0)
            else:
                red, green, blue = red
        
        h, l, s = rgb_to_hls(red, green, blue)
        return (int(round(h * HLSMAX)), int(round(l * HLSMAX)), int(round(s * HLSMAX)))
    except Exception:
        # Return white if anything fails
        return (0, HLSMAX, 0)

def ms_hls_to_rgb(hue, lightness=None, saturation=None):
    """Converts HLSMAX based HLS values to rgb values in the range (0,1)"""
    if lightness is None:
        hue, lightness, saturation = hue
    return hls_to_rgb(hue / HLSMAX, lightness / HLSMAX, saturation / HLSMAX)

def rgb_to_hex(red, green=None, blue=None):
    """Converts (0,1) based RGB values to a hex string 'rrggbb'"""
    if green is None:
        red, green, blue = red
    return ('%02x%02x%02x' % (int(round(red * RGBMAX)), int(round(green * RGBMAX)), int(round(blue * RGBMAX)))).upper()


def get_theme_colors(wb):
    """Gets theme colors from the workbook"""
    # see: https://groups.google.com/forum/#!topic/openpyxl-users/I0k3TfqNLrc
    from openpyxl.xml.functions import QName, fromstring
    colors = []
    if wb.loaded_theme is None:
        return ['FFFFFF']
    
    xlmns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    root = fromstring(wb.loaded_theme)
    themeEl = root.find(QName(xlmns, 'themeElements').text)
    if themeEl is None:
        return ['FFFFFF']
        
    colorSchemes = themeEl.findall(QName(xlmns, 'clrScheme').text)
    if not colorSchemes:
        return ['FFFFFF']
        
    firstColorScheme = colorSchemes[0]
    for c in ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']:
        accent = firstColorScheme.find(QName(xlmns, c).text)
        if accent is None:
            colors.append('FFFFFF')
            continue
            
        for i in list(accent):
            try:
                if 'window' in i.attrib.get('val', ''):
                    # Try to get 'lastClr' attribute, fallback to 'val' or 'FFFFFF'
                    colors.append(i.attrib.get('lastClr', i.attrib.get('val', 'FFFFFF')))
                else:
                    colors.append(i.attrib.get('val', 'FFFFFF'))
            except (KeyError, AttributeError):
                colors.append('FFFFFF')
                
    return colors

def tint_luminance(tint, lum):
    """Tints a HLSMAX based luminance"""
    # See: http://ciintelligence.blogspot.co.uk/2012/02/converting-excel-theme-color-and-tint.html
    if tint < 0:
        return int(round(lum * (1.0 + tint)))
    else:
        return int(round(lum * (1.0 - tint) + (HLSMAX - HLSMAX * (1.0 - tint))))

def theme_and_tint_to_rgb(wb, theme, tint):
    """Given a workbook, a theme number and a tint return a hex based rgb"""
    try:
        # Try to get theme colors safely
        rgb = None
        try:
            rgb = get_theme_colors(wb)[theme]
        except IndexError:
            rgb = get_theme_colors(wb)[0]
            
        # Ensure rgb is a valid hex string
        if not isinstance(rgb, str) or not all(c in '0123456789abcdefABCDEF' for c in rgb):
            return 'FFFFFF'
            
        try:
            h, l, s = rgb_to_ms_hls(rgb)
            return rgb_to_hex(ms_hls_to_rgb(h, tint_luminance(tint, l), s))
        except Exception:
            return 'FFFFFF'
    except Exception:
        # If all else fails, return white
        return 'FFFFFF'