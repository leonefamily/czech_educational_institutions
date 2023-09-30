#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Sep 29 20:25:10 2023

@author: leonefamily
"""

import os
import sys
import time
import geopy
import random
import html5lib  # noqa, ensure parsing html tables with pandas
import argparse
import openpyxl  # noqa, ensure to read xlsx with pandas
import numpy as np
import pandas as pd
import geopandas as gpd
from pathlib import Path
from itertools import product
from selenium import webdriver
from shapely.geometry import Point
from multiprocessing.pool import Pool
from typing import List, Union, Tuple, Dict, Any, Optional

COLUMNS = {
    'Id. zařízení': 'id',
    'Typ': 'type',
    'Název zařízení': 'name',
    'Obec': 'city',
    'Ulice': 'street',
    'Č.p.': 'lrn',
    'Č.o.': 'hn',
    'M.část.': 'city_part',
    'PSČ': 'zip_code',
    'Cizí vyučovací jazyk': 'foreign_lg',
    'Kapacita': 'capacity',
    'Platnost zařízení': 'validity'
}
URL = 'https://rejstriky.msmt.cz/rejskol/default.aspx'
DETAIL = 'VREJVerejne/PravOsoba.aspx'
NORESULT = "Zadaným podmínkám nevyhovuje žádný záznam"
UNI_COUNTS = 'https://dsia.msmt.cz/vystupy/f2/f21.xlsx'
# UNI_SCHOOLS_INFO = 'https://regvssp.msmt.cz/registrvssp/cvslist.aspx'
# FOREIGN_UNI_INFO = 'https://regvssp.msmt.cz/registrvssp/zvssp.aspx'


def get_university_type(
        row: pd.DataFrame
) -> str:
    code = row['code']
    lname = row['name'].lower()
    if code.endswith('000'):
        if code in ['60000', '00000', '10000']:
            return 'indicator'
        return 'university'
    elif 'fakulta' in lname or 'falulta' in lname:  # yes, they misspell
        return 'faculty'
    return 'other'


def get_universities(
        counts_year: int = -1,
        use_gmaps: bool = False,
        gmaps_key: Optional[str] = None
) -> gpd.GeoDataFrame:
    """
    Get data about universities and their locations.

    Parameters
    ----------
    counts_year : int, optional
        Sheet number in data. The default is -1, is last available year.
    use_gmaps : bool, optional
        Use Google Maps Services to geocode. The default is False.
    gmaps_key : Optional[str], optional
        Google Maps API key, if ``use_gmaps``. The default is None.

    Returns
    -------
    gpd.GeoDataFrame

    """
    pre_uni_counts = pd.read_excel(UNI_COUNTS, sheet_name=counts_year)
    uni_counts = pre_uni_counts.iloc[7:-2, np.r_[1:3, 6:17]].dropna()
    uni_counts.columns = [
        'code', 'name', 'total', 'ft_total', 'ft_bach', 'ft_master',
        'ft_fmaster', 'ft_phd', 'dc_total', 'dc_bach', 'dc_master',
        'dc_fmaster', 'dc_phd'
    ]
    types = uni_counts.apply(get_university_type, axis=1)
    uni_counts.insert(1, "type", types)

    last_uni_row = None
    private = False
    for i, row in uni_counts.iterrows():

        if row['code'] == '60000':
            # start of private schools
            private = True

        if last_uni_row is not None and row['type'] == 'faculty':
            uni_counts.loc[i, 'faculty'] = row["name"]
            uni_counts.loc[i, 'university'] = last_uni_row["name"]
            search_string = f'{last_uni_row["name"]}, {row["name"]}'
        elif row['type'] in ['university', 'other']:
            if row['type'] == 'university':
                last_uni_row = row
            else:
                uni_counts.loc[i, 'other'] = row["name"]
            uni_counts.loc[i, 'university'] = last_uni_row["name"]
            search_string = row["name"]
        else:
            search_string = row["name"]

        uni_counts.loc[i, 'private'] = int(private)
        uni_counts.loc[i, 'full_name'] = search_string

    uni_counts.drop(
        uni_counts[uni_counts['type'] == 'indicator'].index,
        inplace=True
    )

    only_names = list(set(n for n in uni_counts['full_name'].tolist() if n))

    places = {
        'full_name': [],
        'address': [],
        'geometry': []
    }
    
    cz = 'Czech Republic'
    default_point = get_locations([cz])[cz]

    if use_gmaps and gmaps_key is not None:
        import googlemaps
        gmaps = googlemaps.Client(key=gmaps_key)
        for name in only_names:
            gcodes = gmaps.geocode(name)
            if gcodes:
                gcode = gcodes[0]
                lon = gcode['geometry']['location']['lng']
                lat = gcode['geometry']['location']['lat']
                places['full_name'].append(name)
                places['address'].append(gcode['formatted_address'])
                places['geometry'].append(Point([lon, lat]))
    else:
        pre_places = get_locations(only_names, keep_orig_response=True)
        for name, place in pre_places.items():
            places['full_name'].append(name)
            places['address'].append('' if place is None else place.address)
            places['geometry'].append(
                default_point if place is None
                else Point([place.longitude, place.latitude])
            )

    places_gdf = gpd.GeoDataFrame(places, crs='epsg:4326')
    universities = gpd.GeoDataFrame(
        uni_counts.merge(places_gdf, how='left', on='full_name'),
        crs='epsg:4326'
    )
    return universities


def glue_address(
        row: pd.Series
) -> str:
    """
    Get string with place address.

    Parameters
    ----------
    row : pd.Series
        Series with used columns.

    Returns
    -------
    str

    """
    addr = row["Obec"] + ','
    addr += '' if pd.isna(row["Ulice"]) else (' ' + row["Ulice"])
    addr += '' if pd.isna(row["Č.p."]) else (' ' + row["Č.p."])
    addr += '' if pd.isna(row["Č.o."]) else ('/' + row["Č.o."])
    return addr


def get_browser(
        headless: bool = True
) -> webdriver.Firefox:
    """
    Get running instance of Firefox browser.

    Parameters
    ----------
    headless : bool, optional
        Whether to hide GUI. The default is True.

    Returns
    -------
    selenium.webdriver.Firefox

    """
    options = webdriver.FirefoxOptions() 
    if headless:
        options.add_argument("-headless") 
    browser = webdriver.Firefox(options=options)
    return browser


def get_locations(
        addresses: List[str],
        keep_orig_response: bool = False
) -> Dict[str, Point]:
    """
    Use Nominatim to get locations of passed addresses or names.
    
    Interval between queries has to be set to 1 second (usage policy).

    Parameters
    ----------
    addresses : List[str]
        List of addresses or places' names.
    keep_orig_response : bool, optional
        Whether to keep Location object as value instead of Point.
        The default is False.

    Returns
    -------
    Dict[str, Point]

    """
    nominatim = geopy.Nominatim(user_agent='schools_placing')
    places = {}
    for i, addr in enumerate(addresses):
        try:
            if addr in places:
                continue
            qstart = time.time()
            coords = nominatim.geocode(addr)
            places[addr] = coords
            qend = time.time()
            qdiff = qend - qstart
            if qdiff < 1:
                time.sleep(1 - qdiff)
            print(round(i * 100 / len(addresses), 2), coords)
        except KeyboardInterrupt:
            raise
        except Exception:
            pass
    if not keep_orig_response:
        geoplaces = {
            addr: Point([place.longitude, place.latitude]) for 
            addr, place in places.items() if place is not None
        }
        return geoplaces
    return places


def merge_locations(
        bigtable: pd.DataFrame
) -> gpd.GeoDataFrame:
    """
    Get addresses locations using Nominatim and merge with existing table.

    Parameters
    ----------
    bigtable : pd.DataFrame
        Existing table of educational institutions.

    Returns
    -------
    gpd.GeoDataFrame

    """
    addresses = list(bigtable['address'].unique())
    geoplaces = get_locations(addresses=addresses)
    geoseries = gpd.GeoSeries(geoplaces).reset_index().rename(
        {'index': 'address', 0: 'geometry'}, axis=1
    )
    bigtable = bigtable.merge(geoseries, on='address')
    schools = gpd.GeoDataFrame(bigtable, crs='epsg:4326')
    return schools


def parse_table_details(
        browser: webdriver.Firefox
) -> pd.DataFrame:
    """
    Extract info table from details of an institution on MŠMT web.

    Requires html5lib module installed for tables parsing.

    Parameters
    ----------
    browser : webdriver.Firefox
        An instance of Firefox browser.

    Returns
    -------
    pd.DataFrame

    """
    try:
        raw_tables = pd.read_html(browser.page_source)
        tabnum = 2 if len(raw_tables) == 3 else 3
        table = raw_tables[tabnum].copy().dropna(how='all')
        table.columns = table.iloc[0]
        table.drop(table.index[0], inplace=True)
    except ValueError:
        table = pd.DataFrame()
    return table


def get_entries(
) -> List[Tuple[str, str]]:
    """
    Get (school_type, location) combinations from MŠMT web using Firefox.

    Returns
    -------
    List[Tuple[str, str]]

    """
    browser = get_browser(headless=True)

    browser.get(URL)
    browser.switch_to.frame(browser.find_element('name', 'mainFrame'))

    regions = [
        el.text for el in
        browser.find_element("name", "ctl39").find_elements("tag name", "option")
        if el.text != ''
    ]
    school_types = [
        el.text for el in
        browser.find_element("name", "ctl38").find_elements("tag name", "option")
        if el.text != ''
    ]
    entries = list(product(school_types, regions))
    return entries


def get_schools(
        entries: List[Tuple[str, str]]
) -> pd.DataFrame:
    """
    Launch an instance of Firefox and parse data for ``entries`` from MŠMT web.

    Parameters
    ----------
    entries : List[Tuple[str, str]]
        List of (school_type, location) combinations.

    Returns
    -------
    pd.DataFrame

    """
    tables = []
    browser = get_browser(headless=True)
    browser.get(URL)
    browser.switch_to.frame(browser.find_element('name', 'mainFrame'))

    for i, (stype, reg) in enumerate(entries):
        try:
            print(f'{stype}, {reg}, {round(i * 100 / len(entries), 2)}%')
            browser.find_element("xpath", f'//option[contains(text(), "{reg}")]').click()
            time.sleep(1)
            browser.find_element("xpath", f'//option[contains(text(), "{stype}")]').click()
            time.sleep(1)
            rownum = browser.find_element("name", 'txtPocetZaznamu')
            rownum.clear()
            rownum.send_keys('9999')
            time.sleep(0.5)
            browser.find_element("name", 'btnVybrat').click()
    
            if NORESULT in browser.page_source:
                browser.close()
                browser = get_browser(headless=True)
                browser.get(URL)
                browser.switch_to.frame(browser.find_element('name', 'mainFrame'))
                continue
    
            details = [el.get_attribute('href') for el
                       in browser.find_elements("tag name", 'a')
                       if DETAIL in el.get_attribute('href')]
    
            for i, detail in enumerate(details):
                browser.get(detail)
                table = parse_table_details(browser)
                tables.append(table)
                browser.execute_script("window.history.go(-1)")
            browser.close()
            browser = get_browser(headless=True)
            browser.get(URL)
            browser.switch_to.frame(browser.find_element('name', 'mainFrame'))
        except Exception as e:
            print(f'Failed {stype}, {reg}: {e}')
    bigtable = pd.concat(tables).drop_duplicates().reset_index(drop=True)
    bigtable['address'] = bigtable.apply(glue_address, axis=1)
    bigtable['address'] = bigtable['address'].replace(
        to_replace =r'Praha \d+', value='Praha', regex=True
    )
    return bigtable


def get_chunks(
        values: List[Any],
        n: int
) -> Tuple[List[List[Any]]]:
    """
    Split a list on a tuple of ``n`` lists with original values.

    Parameters
    ----------
    values : List[Any]
        List with any values
    n : int
        Number of resulting chunks.

    Returns
    -------
    Tuple[List[Any]]

    """
    k, m = divmod(len(values), n)
    return tuple(
        [values[i * k + min(i, m): (i+1) * k + min(i + 1, m)]]
        for i in range(n)
    )


def parse_args(
        args_list: Optional[List[str]] = None
) -> argparse.Namespace:
    if args_list is None:
        args_list = sys.argv[1:]
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-u', '--uni-shp-save-path', required=True,
        help='Where to save geolocalized universities table'
    )
    parser.add_argument(
        '-s', '--sch-shp-save-path', required=True,
        help='Where to save geolocalized universities table'
    )
    parser.add_argument(
        '-c', '--sch-csv-save-path',
        help='Where to save CSV table extract'
    )
    parser.add_argument(
        '-p', '--processes', type=int, default=os.cpu_count(),
        help='Where to save geolocalized table'
    )
    parser.add_argument(
        '-C', '--to-crs',
        help='CRS for saved shapefiles - epsg:XXXX'
    )
    args = parser.parse_args(args_list)
    return args


def main(
        uni_shp_save_path: Union[str, Path],
        sch_shp_save_path: Union[str, Path],
        sch_csv_save_path: Optional[Union[str, Path]] = None,
        processes: int = 35,
        to_crs: Optional[str] = None
):
    """
    Acquire educational institutions data & locations and write them.

    Note, that second stage (after universities) takes an excessive amount
    of time (several hours). Geocoding on Nominatim can be only launched once
    per second according to the rules (so no several processes on this task).
    
    During schools scraping script will launch ``processes`` instances of
    Firefox browser (must be installed and available in the PATH variable).
    Every instance consumes approximately 550MB of RAM, calculate the number of
    ``processes`` depending on available RAM on the machine.

    Parameters
    ----------
    uni_shp_save_path : Union[str, Path]
        Path to save universities data as shapefile.
    sch_shp_save_path : Union[str, Path]
        Path to save schools data as shapefile.
    sch_csv_save_path : Union[str, Path], optional
        Path to save schools data as CSV table. Default is None, no CSV saved.
    processes : int, optional
        How many pseudo-browsers will launch. The default is 25. Use carefully.
    to_crs : int, optional
        Output CRS of shapefiles. The default is 'epsg-4326' (WGS84).

    """
    universities = get_universities()
    if to_crs is not None:
        universities.to_crs(to_crs, inplace=True)
    universities.to_file(
        uni_shp_save_path, encoding='utf-8'
    )

    entries = get_entries()
    random.shuffle(entries)

    chunks = get_chunks(
        values=entries, n=processes
    )

    with Pool() as pool:
        # doesn't work on Windows, function get_schools must be imported from
        # the other file
        tables = pool.starmap(get_schools, chunks)

    bigtable = pd.concat(tables).reset_index(drop=True)
    bigtable.drop_duplicates(inplace=True)
    bigtable.rename(COLUMNS, axis=1, inplace=True)

    if sch_csv_save_path is not None:
        bigtable.to_csv(
            sch_csv_save_path, sep=';', decimal=',', index=False
        )

    schools = merge_locations(bigtable).astype({
        'zip_code': int,
        
    })
    if to_crs is not None:
        schools.to_crs(to_crs, inplace=True)
    schools.to_file(sch_shp_save_path, encoding='utf-8')


if __name__ == '__main__':
    args = parse_args()
    main(
        uni_shp_save_path=args.uni_shp_save_path,
        sch_shp_save_path=args.sch_shp_save_path,
        sch_csv_save_path=args.sch_csv_save_path,
        processes=args.processes,
        to_crs=args.to_crs
    )
