#######################################################  VARIABLES ##################################################################
El.Sub.ID <- 20374
# update working directory
WD <- paste0(
  Working.Folder,"/",El.Sub.ID)
setwd(WD)
#######################################################  PACKAGES ################################################################
library(tidyverse)
library(sf)
library(openxlsx)
library(terra)
library(DBI)
library(odbc)
library(glue)
library(lwgeom)
##################################################  RANGE EXTENT INFERRED ###############################################################
#select layer
BEC13<- st_read(dsn = BEC13.path, layer = "BEC13_v2")
#Queries for Range extent
BGC.Units.Min <- c("BWBSdk", "ESSFdc1", "ESSFdc2", "ESSFmh", "ESSFwm2", "ESSFxc1","ESSFxc2", "ICHdw4", "ICHmk1", "ICHmk2", "ICHmk3", "ICHmw3", "ICHmw5", "ICHvc", "ICHwk1", "ICHxm1", "IDFdk1", "IDFdk2", "IDFdk5", "IDFdm1", "IDFdm2", "IDFxk", "MSdk", "MSdm1", "MSdm2", "MSdm3", "MSdw", "MSxk1", "MSxk2", "SBPSdc", "SBPSmk", "SBPSxc","SBSdh1", "SBSdk", "SBSdw2", "SBSdw3","SBSmc",  "SBSmc2", "SBSmk1", "SBSmk2", "SBSmw", "SBSvk", "SBSwk1")
# minimal evidence supports occurrence of Ws04 in the BWBS. LMH 52 does not list Ws04 in the BWBS or SWB.  LMH68 does not list Ws04 or related Fl05 in any BWBS.  BWBSmw and wk1 seem to be at the NE margin of S. drummondiana and Spiraea douglasii range. Two plots support the occurrence of the Ws04 in the SBSdh1 (Kyla Rushton personal communication)
BGC.Units.Max <- c("BWBSdk","BWBSwk1", "BWBSmw", "ESSFdc1", "ESSFdc2", "ESSFmh", "ESSFwm2", "ESSFxc1","ESSFxc2", "ICHdw4", "ICHmk1", "ICHmk2", "ICHmk3", "ICHmw3", "ICHmw5", "ICHvc", "ICHwk1", "ICHxm1", "IDFdk1", "IDFdk2", "IDFdk5", "IDFdm1", "IDFdm2", "IDFxk", "MSdk", "MSdm1", "MSdm2", "MSdm3", "MSdw", "MSxk1", "MSxk2", "SBPSdc", "SBPSmk", "SBPSxc","SBSdh1", "SBSdk", "SBSdw2", "SBSdw3", "SBSmc", "SBSmc2", "SBSmk1", "SBSmk2", "SBSmw", "SBSvk","SBSwk1", "SBSwk2")

BGC.Range.Min <- BEC13 %>% filter(MAP_LABEL %in% BGC.Units.Min)
BGC.Range.Max <- BEC13 %>% filter(MAP_LABEL %in% BGC.Units.Max)
#MCP around selected BGC.Units
BGC.MCP.Inf.Min <- st_convex_hull(st_union(BGC.Range.Min))
BGC.MCP.Inf.Max <- st_convex_hull(st_union(BGC.Range.Max))
#area calculations
BGC.MCP.Inf.Min.area.km2 <- as.numeric(st_area(BGC.MCP.Inf.Min)) / 1e6
BGC.MCP.Inf.Max.area.km2 <- as.numeric(st_area(BGC.MCP.Inf.Max)) / 1e6
BGC.Range.Min.km2 <- sum(BGC.Range.Min$Shape_Area, na.rm = TRUE)/ 1e6
BGC.Range.Max.km2 <- sum(BGC.Range.Max$Shape_Area, na.rm = TRUE)/ 1e6
################################################## PROJECT BOUNDARIES - Ecosystem Mapping ##########
### Create SF of ALL Ecosystem Mapping Project Boundaries that overlap the widest possible BGC Range
TEI.proj.bound.all <- st_read(
  dsn   = TEI.proj.bound.path,
  query = "
    SELECT *
    FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP
    WHERE PROJECT_TYPE IN (
     'EST', 'EXC', 'NEM', 'NEMNSS', 'NEMSEI', 'NEMWHR',
      'PEM', 'PEMSDM', 'PEMTBT', 'PEMWHR', 'PREM', 'SEI', 'SEIWHR', 'STS',
      'TEM', 'TEMNSS', 'TEMSDM', 'TEMSEI', 'TEMSET', 'TEMTSM', 'TEMWHR',
      'VEG', 'WET'
    )
  "
)
# filter to those that intersect the BGC range
TEI.proj.bound.all <- st_filter(TEI.proj.bound.all, BGC.Range.Max)

###Maxiumum TEI Project Boundary/BGC Range Overlap Estimate
# This excludes PEM and other projects that inventory at a thematic scalle too coarse to capture the 20374
# It includes all TEI datasets that could conceivably detect the 20374, though it is uncertain whether all of these projects effectively inventory for the 20374. This only includes the core BGC Range, not the max range, where occurrence of the 20374 is yet to be confirmed with confidence.
TEI.proj.bound.max <- st_read(
  dsn   = TEI.proj.bound.path,
  query = "
    SELECT *
    FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP
    WHERE BUSINESS_AREA_PROJECT_ID IN (1, 2, 4, 9, 27, 28, 79, 91, 108, 114, 115, 119, 135, 145, 149, 183, 184, 192, 196, 198, 208, 209, 213, 214, 216, 218, 219, 221, 222, 226, 231, 233, 239, 240, 244, 247, 248, 1027, 1037, 1046, 1048, 1049, 1054, 1055, 1058, 1060, 1068, 1070, 4479, 4482, 4498, 4508, 4523, 4917, 5679, 5680, 6130, 6131, 6473, 6484, 6526, 6536, 6537, 6605, 6624)
  ")
overlap_sf <- st_intersection(BGC.Range.Min, TEI.proj.bound.max)
# Dissolve overlapping geometries
overlap_sf <- st_union(overlap_sf)
# Calculate area in m²
overlap_sf$overlap_area_m2 <- st_area(overlap_sf)
# Calculate area of BGC range that is overlapped by selected project boundaries
BGC.Range.Mapped.Max.km2 <- as.numeric(sum(overlap_sf$overlap_area_m2)/ 1e6)
BGC.Range.Mapped.Max.Percent <- as.numeric(BGC.Range.Mapped.Max.km2/BGC.Range.Max.km2)*100
###Minimum TEI Project Boundary Overlap Estimate
# Create SF of selected projects
#TEI.proj.bound.selected <- st_read(dsn   = TEI.proj.bound.path,query = " SELECT * FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP WHERE BUSINESS_AREA_PROJECT_ID IN (4, 135, 216, 233, 244, 1046, 1048, 1049, 1054, 4523, 4917, 6468, 6536, 6624)")
TEI.proj.bound.min <- st_read(
  dsn   = TEI.proj.bound.path,
  query = "
    SELECT *
    FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP
    WHERE BUSINESS_AREA_PROJECT_ID IN (4, 108, 115, 135, 209, 216, 233, 239, 244, 1046, 1048, 1049, 1054, 1055, 4917, 4523, 6473, 6536, 6605, 6624)
  ")
### determine MIN extent of overlap between selected ecosystem mapping projects and BGC Range 
overlap_sf <- st_intersection(BGC.Range.Min, TEI.proj.bound.min)
# Dissolve overlapping geometries
overlap_sf <- st_union(overlap_sf)
# Calculate area in m²
overlap_sf$overlap_area_m2 <- st_area(overlap_sf)
# Calculate area of BGC range that is overlapped by selected project boundaries
BGC.Range.Mapped.Min.km2 <- as.numeric(sum(overlap_sf$overlap_area_m2)/ 1e6)
BGC.Range.Mapped.Min.Percent <- as.numeric(BGC.Range.Mapped.Min.km2/BGC.Range.Min.km2)*100

################################################  BECMASTER - QUERY AND CLEAN ##################################################
# Connect to BECMaster
BEC.Master.Connect <- dbConnect(odbc::odbc(),
                                   .connection_string = paste0("Driver={Microsoft Access Driver (*.mdb, *.accdb)};",
                                                               "DBQ=", BECMaster.path, ";"))
# Query ENV for focal units. Note that this query does not include some plots identified by BECSiteUnit in the Admin table
# valid plots from the Admin table which were not returned by this query were manually added by plot number in the next query
BEC.Master.Focal.Query1 <- "
SELECT 
  e.PlotNumber, e.Date, e.SiteSurveyor, e.Zone, e.SubZone, e.SiteSeries, e.PlotRepresenting, e.Location, 
  e.Longitude, e.Latitude, e.UTMZone, e.UTMEasting, e.UTMNorthing, e.LocationAccuracy, 
  e.SiteModifier1, e.SiteModifier2, e.MapUnit, e.MoistureRegime, e.NutrientRegime, 
  e.StructuralStage, e.StructuralStageMod, e.StandAge, e.Elevation, e.SlopeGradient, e.Aspect, 
  e.MesoSlopePosition, e.Exposure1, e.Exposure2, e.SiteDisturbance1, e.SiteDisturbance2, 
  e.SiteDisturbance3, e.SiteNotes, e.VegSurveyor, e.StrataCoverTree, e.StrataCoverShrub, 
  e.StrataCoverHerb, e.StrataCoverMoss, e.VegNotes, a.BECSiteUnit, a.UserSiteUnit

FROM BECMaster19_Env AS e
LEFT JOIN BECMaster19_Admin AS a
  ON e.PlotNumber = a.Plot
WHERE e.SiteSeries LIKE '%Ws04%'
"
BEC.Master.Focal1 <- dbGetQuery(BEC.Master.Connect, BEC.Master.Focal.Query1)
BEC.Master.Focal.Query2 <- "
SELECT 
  e.PlotNumber, e.Date, e.SiteSurveyor, e.Zone, e.SubZone, e.SiteSeries, e.PlotRepresenting, e.Location, 
  e.Longitude, e.Latitude, e.UTMZone, e.UTMEasting, e.UTMNorthing, e.LocationAccuracy, 
  e.SiteModifier1, e.SiteModifier2, e.MapUnit, e.MoistureRegime, e.NutrientRegime, 
  e.StructuralStage, e.StructuralStageMod, e.StandAge, e.Elevation, e.SlopeGradient, e.Aspect, 
  e.MesoSlopePosition, e.Exposure1, e.Exposure2, e.SiteDisturbance1, e.SiteDisturbance2, 
  e.SiteDisturbance3, e.SiteNotes, e.VegSurveyor, e.StrataCoverTree, e.StrataCoverShrub, 
  e.StrataCoverHerb, e.StrataCoverMoss, e.VegNotes, a.BECSiteUnit, a.UserSiteUnit

FROM BECMaster19_Env AS e
LEFT JOIN BECMaster19_Admin AS a
  ON e.PlotNumber = a.Plot
WHERE a.BECSiteUnit LIKE '%Ws04%' OR a.UserSiteUnit LIKE '%Ws04%'
"
BEC.Master.Focal2 <- dbGetQuery(BEC.Master.Connect, BEC.Master.Focal.Query2)
# bind BECmaster queries from ENV and ADMIN tables, keep only unique records.
BEC.Master.Focal <-unique(rbind(BEC.Master.Focal1,BEC.Master.Focal2))
# Force longitudes to be negative, lat to be numeric
BEC.Master.Focal$Longitude <- -abs(as.numeric(BEC.Master.Focal$Longitude))
BEC.Master.Focal$Latitude  <- as.numeric(BEC.Master.Focal$Latitude)

# list of plots vetted either by BECSiteUnit field or by Kristi Iverson email 20251125
BEC.Vetted <- c(
  '14-3231','14-5469','272980','378480','8400061','9613288','9613857','9615182','9615223',
  '9617072','9628709','9628714','9628718','9628877','9628878','9628880','963106','9650957',
  'CL00015','CL00035','CL00041','CL00061','CL00062','CL00063','N05-182'
)
# Add new field 'Include' to filter vetted plots 
BEC.Master.Focal <- BEC.Master.Focal %>%
  mutate(
    Include = dplyr::if_else(
      PlotNumber %in% BEC.Vetted, # The logical test
      "Yes",                      # Value if TRUE
      "No"                        # Value if FALSE
    )
  )


# REMOVES ROWS WITH NO COORDS. Convert to sf object using Longitude and Latitude
BEC.Master.Focal.sf <- BEC.Master.Focal %>%
  filter(is.finite(Longitude), is.finite(Latitude)) %>%
  st_as_sf(coords = c("Longitude", "Latitude"), crs = 4326)
# Reproject to albers
BEC.Master.Focal.sf <- st_transform(BEC.Master.Focal.sf, crs = 3005)
# Buffer each point by half the LocationAccuracy for viewing in GIS
BEC.Master.Focal.sf.buffer <- BEC.Master.Focal.sf %>%
  mutate(geometry = st_buffer(geometry, dist = LocationAccuracy / 2))
################################################  OTHER PLOTS - QUERY AND CLEAN ##################################################
# These are the WIlliston Wetlands plots from BAPID 6538
plots.6538.2017 <- st_read(dsn   = "V:/ESR Spatial/20374/data/BAPID 6538 material from EcoCat/FieldData/FieldData_2017_raw.shp",  query = "SELECT * FROM FieldData_2017_raw WHERE WetlandAss = 'Ws04'")
plots.6538.2019 <- st_read(dsn  = "V:/ESR Spatial/20374/data/BAPID 6538 material from EcoCat/FieldData/FieldData_2019_raw.shp", query = "SELECT * FROM FieldData_2019_raw WHERE WetlandAss = 'Ws04'")
# Map 2019 PlotNum -> PlotNo 
if ("PlotNum" %in% names(plots.6538.2019) && !"PlotNo" %in% names(plots.6538.2019)) {
  plots.6538.2019 <- plots.6538.2019 %>%
    mutate(PlotNo = PlotNum)
}
# Keep only common columns and bind
common <- intersect(names(plots.6538.2017), names(plots.6538.2019))
plots.6538.2017x <- plots.6538.2017[, common]
plots.6538.2019x <- plots.6538.2019[, common]
plots.6538 <- rbind(plots.6538.2017x, plots.6538.2019x)
# Buffer for compatibility with NOO analysis. Use 30 m buffer
plots.6538 <- plots.6538 %>%
  mutate(geometry = st_buffer(`_ogr_geometry_`, dist = 30))
################################################# TEI - QUERY AND CLEAN #########################################################
### TEM ###
# These queries take about 1 hour to run
# These queries warnings that don't seem to matter; "In CPL_read_ogr(dsn, layer, query, as.character(options), quiet,  :GDAL Error 1: Error occurred in filegdbtable.cpp at line 1777"
### Started with a wildcard query (~LIKE '%Ws04%') not required as only clean values (Ws04) returned 20251211
# do not need to query SITEMC fields as they return no values, but note that they do in 6473 and 6605 which are manually entered below
Query.TEM.s <- r"(
SELECT
  SDEC_1, SDEC_2, SDEC_3,
  SITE_S1, SITE_S2, SITE_S3,
  STRCT_S1, STRCT_S2, STRCT_S3,
  STAND_A1, STAND_A2, STAND_A3,
  SITEMC_S1, SITEMC_S2, SITEMC_S3,
  BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE,
  PLOT_NO, POLY_COM,
  PROJ_ID, PROJ_TYPE, OBJECTID, TEIS_ID, PROJPOLYID, BAPID,
  Shape_Area, shape
FROM TEIS_Master_Long_Tbl
WHERE PROJ_TYPE IN (
  'ESA','EST','EXC','NEM','NEMNSS','NEMSEI','NEMWHR',
  'PEM','PEMSDM','PEMTBT','PEMWHR',
  'PREM','SEI','SEIWHR','STS',
  'TEM','TEMNSS','TEMSDM','TEMSEI','TEMSET','TEMTSM','TEMWHR',
  'VEG','WET'
)
AND (SITE_S1 = 'Ws04' OR SITE_S2 = 'Ws04' OR SITE_S3 = 'Ws04')
)"


TEM.s <- st_read(
  dsn = TEI.Path,
  query = Query.TEM.s,
  quiet = TRUE
)

# Create a new column that concatenates the four BGC fields
TEM.s <- TEM.s %>%
  unite(
    col = BGC.Full,  # New column name
    BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE, # Columns to combine
    sep = "",
    na.rm = TRUE, #not insert NA for NULL values
    remove = FALSE
  )

### Mapcode Query
# Initial query - narrows to relevant BAPIDS
Query.TEM.m <- "
SELECT 
    SDEC_1, SDEC_2, SDEC_3,
    SITE_S1, SITE_S2, SITE_S3,
    STRCT_S1, STRCT_S2, STRCT_S3,
    STAND_A1, STAND_A2, STAND_A3,
    SITEMC_S1, SITEMC_S2, SITEMC_S3,
    BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE,
    PLOT_NO, POLY_COM,
    PROJ_ID, PROJ_TYPE, OBJECTID, TEIS_ID, PROJPOLYID, BAPID,
    Shape_Area, shape
FROM TEIS_Master_Long_Tbl
WHERE BAPID IN (4, 108, 115, 135, 209, 216, 233, 239, 244, 1046, 1048, 1049, 1054, 1055)
"

TEM.m <- st_read(
  dsn   = TEI.Path,
  query = Query.TEM.m,
  quiet = TRUE
)

# concatenate BGC
TEM.m <- TEM.m %>%
  unite(
    col = BGC.Full,  # New column name
    BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE, # Columns to combine
    sep = "",
    na.rm = TRUE, #not insert NA for NULL values
    remove = FALSE
  )

# I checked and map codes are not appearing in site code field.
# Further narrows to relevant mapcodes
fields <- c("SITEMC_S1","SITEMC_S2","SITEMC_S3")
vals <- c("DS", "WS", "WD", "DM")

TEM.m <- TEM.m %>%
  filter(if_any(all_of(fields), ~ . %in% vals))
# Further narrows to relevant BGCs
TEM.m <- TEM.m %>%
  filter(BGC.Full %in% c(
    "ESSFdc2", "ESSFxc", "IDFdk2", "IDFdm1", "MSdm1", 
    "MSdm2", "MSxk", "SBPSmk", "SBPSxc", "SBSdk", 
    "SBSdw2", "SBSmc2", "SBSmw", "SBSwk1"
  ))

# Final query for explicit BGC/BAPID/Mapcodes
filter_condition <- 
  (TEM.m$BGC.Full == 'SBSdk' & TEM.m$BAPID == 4 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'IDFdk2' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'MSxk' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'MSdm2' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 135 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'SBSdk' & TEM.m$BAPID == 216 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmc2' & TEM.m$BAPID == 216 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 233 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 244 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'SBSwk1' & TEM.m$BAPID == 1046 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBSdw2' & TEM.m$BAPID == 1048 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBPSmk' & TEM.m$BAPID == 1049 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBPSxc' & TEM.m$BAPID == 1054 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'IDFdm1' & TEM.m$BAPID == 1055 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSdk' & TEM.m$BAPID == 115 & (TEM.m$SITEMC_S1 == 'DM' | TEM.m$SITEMC_S2 == 'DM' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmc2' & TEM.m$BAPID == 115 & (TEM.m$SITEMC_S1 == 'DM' | TEM.m$SITEMC_S2 == 'DM' | TEM.m$SITEMC_S3 == 'DM')) 

TEM.m <- TEM.m %>%
  filter(!!filter_condition)

### Manually add BAPID 6605
Query.TEM.6605 <- c(
  "SELECT",
  "SDEC_1, SDEC_2, SDEC_3,",
  "SITE_S1, SITE_S2, SITE_S3,",
  "STRCT_S1, STRCT_S2, STRCT_S3,",
  "STAND_A1, STAND_A2, STAND_A3,",
  "SITEMC_S1, SITEMC_S2, SITEMC_S3, BGC_VLD,",
  "PLOT_NO, POLY_COM,",
  "PROJ_ID, PROJ_TYPE, OBJECTID, TEIS_ID, PROJPOLYID, BAPID,",
  "Shape_Area, shape",
  "FROM TEI_Long_Tbl",
  "WHERE",
  " SITE_S1 = 'Ws04' OR",
  " SITE_S2 = 'Ws04' OR",
  " SITE_S3 = 'Ws04'"
)

Query.TEM.6605 <- paste(Query.TEM.6605, collapse = " ")

TEM.6605 <- st_read(
  dsn   = TEI_BAPID6605_path,
  query = Query.TEM.6605,
  quiet = TRUE
)

TEM.6605 <- TEM.6605 %>%
  rename(BGC.Full = BGC_VLD)

### Manually add BAPID 6473
Query.TEM.6473 <- c(
  "SELECT",
  "SDEC_1, SDEC_2, SDEC_3,",
  "SITE_S1, SITE_S2, SITE_S3,",
  "STRCT_S1, STRCT_S2, STRCT_S3,",
  "STAND_A1, STAND_A2, STAND_A3,",
  "SITEMC_S1, SITEMC_S2, SITEMC_S3,",
  "BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE,",
  "PLOT_NO, POLY_COM,",
  "PROJ_ID, PROJ_TYPE, OBJECTID, TEIS_ID, PROJPOLYID, BAPID,",
  "Shape_Area, shape",
  "FROM TEI_Long_Tbl_6473",
  "WHERE",
  " SITEMC_S1 = 'Ws04' OR",
  " SITEMC_S2 = 'Ws04' OR",
  " SITEMC_S3 = 'Ws04'")

Query.TEM.6473 <- paste(Query.TEM.6473, collapse = " ")

TEM.6473 <- st_read(
  dsn   = TEI_BAPID6473_path,
  query = Query.TEM.6473,
  quiet = TRUE
)

TEM.6473 <- TEM.6473 %>%
  unite(
    col = BGC.Full,  # New column name
    BGC_ZONE, BGC_SUBZON, BGC_VRT, BGC_PHASE, # Columns to combine
    sep = "",
    na.rm = TRUE, #not insert NA for NULL values
    remove = FALSE
  )

# bind tem from site code queries with TEM from mapcode query
TEM <- bind_rows(TEM.s, TEM.m, TEM.6473, TEM.6605)
# remove bad records based on manual review
# remove records where structual stage is not LIKE %3.  Manually reviewed these records with imagery. Most were not associated with a stream or had burned. There was also a decile error in one
bad.teis <- c(22905043, 23666961, 23667093, 23667543, 25411952, 25412351, 25412356, 25412570, 25412598)
TEM <- TEM %>% filter(!TEIS_ID %in% bad.teis)

# add lat/long fields
TEM <- TEM %>%
  bind_cols(
    TEM %>% 
      st_centroid() %>% # get centroid
      st_transform(4326) %>%  # change Alberts to geogrpahic 
      st_coordinates() %>% # extract coordinates
      as.data.frame() %>% # add columns
      rename(Latitude = Y, Longitude = X)
  )
### Fill in missing BGC based on manual review
TEM <- TEM %>%
  mutate(
    BGC.Full = case_when(
      OBJECTID %in% c(3015989) ~ "SBSdw2",
      OBJECTID %in% c(3016307, 3016308) ~ "ICHmk3",
      TRUE ~ BGC.Full
    )
  )
######################################################## AOO Sum from TEI #######################################################
### Add AOO field to TEM to multiply deciles for focal site units.
TEM <- TEM %>%
  mutate(
    AOO = case_when(
      # Conditions based on SITE
      SITE_S1 == 'Ws04' ~ SDEC_1 * 0.1 * Shape_Area,
      SITE_S2 == 'Ws04' ~ SDEC_2 * 0.1 * Shape_Area,
      SITE_S3 == 'Ws04' ~ SDEC_3 * 0.1 * Shape_Area,
      SITEMC_S1 == 'Ws04' ~ SDEC_1 * 0.1 * Shape_Area,
      SITEMC_S2 == 'Ws04' ~ SDEC_2 * 0.1 * Shape_Area,
      SITEMC_S3 == 'Ws04' ~ SDEC_3 * 0.1 * Shape_Area,
      
      #Conditions based on BGC.Full, BAPID, and SITEMC
      
      # SBSdk / WS (BAPID 4 and 216)
      BGC.Full == 'SBSdk' & BAPID %in% c(4,216) & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdk' & BAPID %in% c(4,216) & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdk' & BAPID %in% c(4,216) & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBSmw / WD (BAPID 135, 233, 244)
      BGC.Full == 'SBSmw' & BAPID %in% c(135, 233, 244) & SITEMC_S1 == 'WD' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSmw' & BAPID %in% c(135, 233, 244) & SITEMC_S2 == 'WD' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSmw' & BAPID %in% c(135, 233, 244) & SITEMC_S3 == 'WD' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBSmc2 / WS (BAPID 216)
      BGC.Full == 'SBSmc2' & BAPID == 216 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSmc2' & BAPID == 216 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSmc2' & BAPID == 216 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBSwk1 / DS (BAPID 1046)
      BGC.Full == 'SBSwk1' & BAPID == 1046 & SITEMC_S1 == 'DS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSwk1' & BAPID == 1046 & SITEMC_S2 == 'DS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSwk1' & BAPID == 1046 & SITEMC_S3 == 'DS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBSdw2 / DS (BAPID 1048)
      BGC.Full == 'SBSdw2' & BAPID == 1048 & SITEMC_S1 == 'DS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdw2' & BAPID == 1048 & SITEMC_S2 == 'DS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdw2' & BAPID == 1048 & SITEMC_S3 == 'DS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBPSmk / DS (BAPID 1049)
      BGC.Full == 'SBPSmk' & BAPID == 1049 & SITEMC_S1 == 'DS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBPSmk' & BAPID == 1049 & SITEMC_S2 == 'DS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBPSmk' & BAPID == 1049 & SITEMC_S3 == 'DS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # SBPSxc / DS (BAPID 1054)
      BGC.Full == 'SBPSxc' & BAPID == 1054 & SITEMC_S1 == 'DS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBPSxc' & BAPID == 1054 & SITEMC_S2 == 'DS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBPSxc' & BAPID == 1054 & SITEMC_S3 == 'DS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # IDFdk2 / WS (BAPID 108)
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # IDFdm1 / WS (BAPID 1055)
      BGC.Full == 'IDFdm1' & BAPID %in% c(1055) & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdm1' & BAPID %in% c(1055) & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdm1' & BAPID %in% c(1055) & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSxk / WS (BAPID 108)
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSdm2 / WS (BAPID 108)
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSdm2 / DM (BAPID 115)
      BGC.Full %in% c('SBSmc2', 'SBSdk') & BAPID == 115 & SITEMC_S1 == 'DM' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full %in% c('SBSmc2', 'SBSdk') & BAPID == 115 & SITEMC_S2 == 'DM' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full %in% c('SBSmc2', 'SBSdk') & BAPID == 115 & SITEMC_S3 == 'DM' ~ SDEC_3 * 0.1 * Shape_Area,
      # Default case
      TRUE ~ 0.0 
    )
  )
# add Km2 field
TEM <-TEM %>%
  mutate(AOO.km2 = TEM$AOO/1e6)

# Sum AOO Max
AOO.Obs.Max.km2 <- sum(as.numeric(TEM$AOO),na.rm = TRUE)/1e6
# Sum AOO Min based on removal of BWBS and SBSwk2. See comments under Range Extent
AOO.Obs.Min.km2 <- sum(as.numeric(TEM$AOO[TEM$BGC_ZONE != 'BWBS' & TEM$BGC_ZONE != 'SBSwk2']), na.rm = TRUE) / 1e6

### Determination of spatial pattern = "Small Patch"
Patch.Size.Avg.ha <- mean(TEM$AOO, na.rm = TRUE) / 1e4 # 5 h
Patch.Size.Med.ha <- median(TEM$AOO, na.rm = TRUE) / 1e4 # 2 h
Patch.Size.Max.ha <- max(TEM$AOO, na.rm = TRUE) / 1e4 # 181 ha
################################################  AOO Estimate ################################################
## Five TEM Projects (BAPID numbers: 4, 216, 4523, 4917, 6536,6605) recorded occupancy rates for this ecological community across five BGC units (ESSFdc2, IDFdk2, SBSdk, SBSmc2, SBSmk1). The rate of occupancy ranged from 0.03% to 1.02% (excluding obvious outliers , or instances where the mapping was geographically constrained in a manner that would bias occupancy rates). These rates of occupancy are multiplied by the min and max BGC range estimates to calculate an estimated area of occupancy below.
# BGC.Range.Mapped.Percent and BGC.Range.Mapped.km2 are defined above
AOO.Est.Min.km2 <- (BGC.Range.Min.km2*0.0003)
AOO.Est.Max.km2 <- (BGC.Range.Max.km2*0.0102)
#
#####################################################  NOO  #################################################
# Parameters
sep.dist.m <- 1000 #m
min.occ.size.km2 <- 0.0005 # 2 ha
### Create a new object of all occurrences merging features from BEC.Master.Focal.sf and TEM. Note that unvetted BEC plots are filtered at this stage, but further filtering of TEM occurs in the NOO process.
# Filter to vetted plots
BEC.Vetted.Filtered <- BEC.Master.Focal.sf.buffer %>% 
  filter(Include == "Yes")
#Force name to 'geometry' and cast to Multipolygon
BEC_prep <- BEC.Vetted.Filtered %>%
  select(-Include) %>%
  st_set_geometry("geometry") %>%
  st_cast("MULTIPOLYGON")
#Force name to 'geometry' and cast to Multipolygon
# st_set_geometry() is used twice here: once to rename 'shape' to 'geometry' and once to ensure it is the active spatial column
TEM_prep <- TEM %>%
  st_make_valid() %>% 
  rename(geometry = shape) %>% 
  st_set_geometry("geometry") %>% 
  st_cast("MULTIPOLYGON", warn = FALSE)
#Force name to 'geometry' and cast to Multipolygon
plots.6538_prep <- plots.6538 %>%
  select(-Aspect) %>%
  st_set_geometry("geometry") %>%
  st_cast("MULTIPOLYGON")
# bind TEM and BEC Plots 
# When creating Occurrences, ensure all objects are verified sf objects before binding
# bind_rows can sometimes drop the sf class if the first object is not a clean sf object
Occurrences <- bind_rows(
  st_as_sf(BEC_prep), 
  st_as_sf(TEM_prep), 
  st_as_sf(plots.6538_prep)
)
# Set the active geometry explicitly
Occurrences <- st_set_geometry(Occurrences, "geometry")
# Reset metadata for Terra
Occurrences <- st_as_sf(Occurrences)
### minimum NOO estimate. Removing BWBS and SBSwk2 and NA values from the bind
valid.occurrences.min <- Occurrences %>% 
  filter(
    (BGC.Full != 'BWBSmk' & BGC.Full != 'BWBSwk1' & BGC.Full != 'SBSwk2') | 
      is.na(BGC_ZONE) | 
      is.na(BGC.Full)
  )
# Keep only BEC Plots and TEM that meets minimum size
min_occ_clean <- valid.occurrences.min %>% 
  filter(AOO.km2 >= min.occ.size.km2 | is.na(AOO.km2)) %>%
  filter(!st_is_empty(.))
# remove empty records
min_occ_clean <- min_occ_clean[!st_is_empty(min_occ_clean), ]
# convert to terra object
min_occ_vect <- vect(min_occ_clean)
# create buffers
occ_buf <- buffer(min_occ_vect, width = (sep.dist.m/2))
# intersect buffers
adj_mat <- relate(occ_buf, occ_buf, relation = "intersects")
# Group buffers based on adjacency
n <- nrow(adj_mat)
group_id <- rep(NA_integer_, n)
current_group <- 0L

for (i in seq_len(n)) {
  if (is.na(group_id[i])) {
    current_group <- current_group + 1L
    to_visit <- i
    while (length(to_visit) > 0) {
      idx <- to_visit[1]
      to_visit <- to_visit[-1]
      if (is.na(group_id[idx])) {
        group_id[idx] <- current_group
        neighbors <- which(adj_mat[idx, ])
        if (length(neighbors)) {
          to_visit <- unique(c(to_visit, neighbors))
        }
      }
    }
  }
}
# Attach group IDs to buffers
occ_buf$GroupID <- group_id
# Dissolve buffers by GroupID
NOO.Min <- aggregate(occ_buf, by = "GroupID", dissolve = TRUE)
# convert back to sf
NOO.Min <- st_as_sf(NOO.Min)
### maximum NOO estimate including the BWBS and SBSwk2. See comments in Range extent. ###
# Keep only BEC Plots and TEM that meets minimum size
max_occ_clean <- Occurrences %>% 
  filter(AOO.km2 >= min.occ.size.km2 | is.na(AOO.km2)) %>%
  filter(!st_is_empty(.))
# remove empty records
max_occ_clean <- max_occ_clean[!st_is_empty(max_occ_clean), ]
# convert to terra object
max_occ_vect <- vect(max_occ_clean)
# create buffers
occ_buf <- buffer(max_occ_vect, width = (sep.dist.m/2))
# intersect buffers
adj_mat <- relate(occ_buf, occ_buf, relation = "intersects")
# Group buffers based on adjacency
n <- nrow(adj_mat)
group_id <- rep(NA_integer_, n)
current_group <- 0L

for (i in seq_len(n)) {
  if (is.na(group_id[i])) {
    current_group <- current_group + 1L
    to_visit <- i
    while (length(to_visit) > 0) {
      idx <- to_visit[1]
      to_visit <- to_visit[-1]
      if (is.na(group_id[idx])) {
        group_id[idx] <- current_group
        neighbors <- which(adj_mat[idx, ])
        if (length(neighbors)) {
          to_visit <- unique(c(to_visit, neighbors))
        }
      }
    }
  }
}

# Attach group IDs to buffers
occ_buf$GroupID <- group_id
# Dissolve buffers by GroupID
NOO.Max <- aggregate(occ_buf, by = "GroupID", dissolve = TRUE)
# convert back to sf
NOO.Max <- st_as_sf(NOO.Max)

### Join group ID to Occurrences 
# 1. Join Min Group IDs
Occurrences <- Occurrences %>%
  st_join(NOO.Min %>% select(GroupID_Min = GroupID), join = st_intersects, largest = TRUE)

# 2. Join Max Group IDs
Occurrences <- Occurrences %>%
  st_join(NOO.Max %>% select(GroupID_Max = GroupID), join = st_intersects, largest = TRUE)
#### Count NOO
NOO.Count.Min <- nrow(NOO.Min)
NOO.Count.Max <- nrow(NOO.Max)
##################################################  RANGE EXTENT OBSERVED ###############################################################
### Take object from NOO (combined TEM and BEC.Master.Focal.sf) and use it to create a minimum convex polygon around all recorded occurrences
BGC.MCP.Obs.Min <- st_convex_hull(st_union(Occurrences))
BGC.MCP.Obs.Min.area.km2 <- as.numeric(st_area(BGC.MCP.Obs.Min)) / 1e6
#####################################################  IMPACTS #################################################
Human.Disturbance <- st_read(My.CEF.Human.Disturbance.path, "BC_CEF_HUMAN_DISTURBANCE_2023_fixed")

# Join (left join keeps all TEM rows).attaches all HD columns to TEM where they intersect.
Impacts <- st_join(
  TEM,
  Human.Disturbance,
  join = st_intersects,
  left = TRUE
)
# summarize impacts
impact.summary.df <- Impacts %>%
  st_drop_geometry() %>%                    
  group_by(CEF_DISTURB_GROUP) %>%          
  summarize(
    unique_teis_count = n_distinct(TEIS_ID, na.rm = TRUE)
  ) %>%
  mutate(
    CEF_DISTURB_GROUP = ifelse(is.na(CEF_DISTURB_GROUP), "No Recorded Human Disturbance", CEF_DISTURB_GROUP),
    percent = round((unique_teis_count / nrow(TEM)) * 100, 0)
  )
#####################################################  THREATS ################################################# 
### EXPERIMENTAL (not applied in this CSA)
### Determine percentage of AOO polygons that are in the proxy THLB.
#bounding box around max MCP to narrow down initial pTHLB query
#bbox.wkt <- st_as_text(st_as_sfc(st_bbox(BGC.MCP.Inf.Max)))
# Proxy THLB (pTHLB) is like THLB but calculated more regularly at the provincial scale.
#THLB Proxy GDB query uses bounding box wkt filter 
#pTHLB <- st_read(dsn = THLB.Proxy.Path, 
                 #layer = "provincial_pthlb", 
                 #wkt_filter = bbox.wkt)
### This is the most efficient 'select by location' process that I know.
### Selects (but not clip) THLB polygons to Max BGC Range
# Simplifies sf math from 3D to 2D
#sf_use_s2(FALSE)
# creates a spatial index of the THLB data. Each feature in the list recieves a number identifying a BGC polygon if there is overlap
#matches <- st_intersects(pTHLB, Occurrences)
# creates a TRUE/FALSE vector from matches 
#keep_idx <- lengths(matches) > 0
# selects every row (and all columns) for records based in the keep_idx Logical vector = TRUE
#result <- pTHLB[keep_idx, ]
# dissolved (union) Occurrences and use to crop (intersection pTHLB)
#pthlb_cropped <- st_intersection(result, st_union(Occurrences))
# Re-calculate area on the cropped pTHLB
#pthlb_cropped$area_cropped <- st_area(pthlb_cropped)
# Drop geometry
#thlb_df <- st_drop_geometry(pthlb_cropped)

# Calculate the percentage of total pTHLB polygons in occurrence polygons that are pTHLB
#thlb_percent <- (sum(as.numeric(thlb_df$area_cropped) * thlb_df$pthlb_fact, na.rm = TRUE)/sum(as.numeric(thlb_df$area_cropped), na.rm = TRUE)) * 100
# Cleanup memory
#rm(pTHLB, result)
#gc()
#sf_use_s2(TRUE)
#st_write(pthlb_cropped, "pthlb.gpkg",layer = "pthlb")
#################################################### WRITE SPATIAL #################################################################
# define output path
out.gpkg <- file.path(getwd(), "outputs", paste0(El.Sub.ID, "-csa-", format(Sys.time(), "%Y%m%d_%H%M"), ".gpkg"))

#write BEC Master points GPKG
st_write(BEC.Master.Focal.sf.buffer, out.gpkg, layer = "BECMaster (Buffered Locational Uncertainty)")
#write TEM
st_write(TEM, out.gpkg, layer = "TEM (all)")
# Write Range
st_write(BGC.Range.Min, out.gpkg, layer ="BGC Mapped Area Min")
st_write(BGC.Range.Max, out.gpkg, layer ="BGC Mapped Area Max")
st_write(BGC.MCP.Inf.Min, out.gpkg, layer = "Range Extent Inferred Area Min (MCP)")
st_write(BGC.MCP.Inf.Max, out.gpkg, layer = "Range Extent Inferred Area Max (MCP)")
st_write(BGC.MCP.Obs.Min, out.gpkg, layer = "Range Extent Observed Area Min (MCP)")
# Write NOO
st_write(NOO.Max, out.gpkg, layer = "NOO Max Clusters")
st_write(NOO.Min, out.gpkg, layer = "NOO Min Clusters")
# Write Occurrences
st_write(Occurrences, out.gpkg, layer = "Occurrences (Selected TEM and BECMaster records")
# Write selected project boundaries and write ALL project boundaries (for ease of review)
st_write(TEI.proj.bound.all, out.gpkg, layer = "TEI Project Boundaries (All)")
st_write(TEI.proj.bound.min, out.gpkg, layer = "TEI Project Boundaries (Min - valid)")
st_write(TEI.proj.bound.max, out.gpkg, layer = "TEI Project Boundaries (Max - valid)")
############################################################### WRITE TABLES ####################################################
### Round all values for reporting
round <- list(
  # Range
  BGC.MCP.Inf.Min.area.km2     = round(BGC.MCP.Inf.Min.area.km2, 0),
  BGC.MCP.Inf.Max.area.km2     = round(BGC.MCP.Inf.Max.area.km2, 0),
  BGC.MCP.Obs.Min.area.km2     = round(BGC.MCP.Obs.Min.area.km2, 0),
  BGC.Range.Min.km2            = round(BGC.Range.Min.km2, 0),
  BGC.Range.Max.km2            = round(BGC.Range.Max.km2, 0),
  BGC.Range.Mapped.Min.Percent = round(BGC.Range.Mapped.Min.Percent, 1),
  BGC.Range.Mapped.Max.Percent = round(BGC.Range.Mapped.Max.Percent, 1),
  BGC.Range.Mapped.Min.km2     = round(BGC.Range.Mapped.Min.km2, 1),
  BGC.Range.Mapped.Max.km2     = round(BGC.Range.Mapped.Max.km2, 1),
  # AOO
  AOO.Obs.Min.km2              = round(AOO.Obs.Min.km2, 0),
  AOO.Obs.Max.km2              = round(AOO.Obs.Max.km2, 0),
  AOO.Est.Min.km2              = round(AOO.Est.Min.km2, 0),
  AOO.Est.Max.km2              = round(AOO.Est.Max.km2, 0),
  Patch.Size.Avg.ha            = round(Patch.Size.Avg.ha, 1),
  Patch.Size.Med.ha            = round(Patch.Size.Med.ha, 1),
  Patch.Size.Max.ha            = round(Patch.Size.Max.ha, 1),
  # NOO
  NOO.Count.Min                = round(NOO.Count.Min, 0),
  NOO.Count.Max                = round(NOO.Count.Max, 0),
  min.occ.size.km2             = min.occ.size.km2, 
  sep.dist.m                   = sep.dist.m        
)
### Create Excel doc to begin writing
data.xlsx <- createWorkbook()
### WRITE DATA ###
### BEC Master Data 
addWorksheet(data.xlsx, "BECMaster")
writeDataTable(data.xlsx, sheet = "BECMaster", x = BEC.Master.Focal, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")
setColWidths(data.xlsx, sheet = "BECMaster", cols = 1:50, widths = "auto")
### Project Boundaries Data 
TEI.df <- st_drop_geometry(TEI.proj.bound.all)
addWorksheet(data.xlsx, "TEI Projects")
writeDataTable(data.xlsx, sheet = "TEI Projects", x = TEI.df, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")
setColWidths(data.xlsx, sheet = "TEI Projects", cols = 1:50, widths = "auto")
### TEM Data 
TEM.df <- st_drop_geometry(TEM)
addWorksheet(data.xlsx, "TEI")
writeDataTable(data.xlsx, sheet = "TEI", x = TEM.df, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")
setColWidths(data.xlsx, sheet = "TEI", cols = 1:50, widths = "auto")
### Impacts Data
Impacts.df <- st_drop_geometry(Impacts)
addWorksheet(data.xlsx, "Impact (raw)")
writeDataTable(data.xlsx, sheet = "Impact (raw)", x = Impacts.df, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")
setColWidths(data.xlsx, sheet = "Impact (raw)", cols = 1:50, widths = "auto")
### Geographic Data
#BEC Master formatting
# add fields
BEC.Geographic.Summary <- BEC.Master.Focal %>%
  mutate(
    `Element Subnational ID` = El.Sub.ID,            # constant value in every row
    ID                      = PlotNumber,            # from BEC.Master.Focal$PlotNumber
    Source                  = "BEC_Master",          # constant string for all rows
    Notes                   = ""                     # blank notes
  ) %>%
# select and sort
  select(
    `Element Subnational ID`,
    ID,
    Latitude,
    Longitude,
    Source,
    Include,
    Notes
  )

#TEI formatting
TEI.Geographic.Summary <- TEM %>%
  # Select the attributes and rename/recode them to match Geographic.Summary
  mutate(
    `Element Subnational ID` = El.Sub.ID, # Match the constant from the original set
    ID                       = as.character(TEIS_ID),   # Use TEIS_ID for the new 'ID' column
    Source                   = 'TEI',     # Constant string for the new source
    Include                  = 'Yes',     # Constant 'Yes'
    Notes                    = "",        # Blank notes
  ) %>%
  # Select and sort
  select(
    `Element Subnational ID`,
    ID,
    Latitude,
    Longitude,
    Source,
    Include,
    Notes
  ) %>%
# drop geometry and Append
st_drop_geometry()
Geographic.Summary <- bind_rows(BEC.Geographic.Summary,TEI.Geographic.Summary)

# Write to Excel as a formatted table
addWorksheet(data.xlsx, "Geographic Summary")

writeDataTable(
  wb         = data.xlsx,
  sheet      = "Geographic Summary",
  x          = Geographic.Summary,
  startCol   = 1,
  startRow   = 1,
  tableStyle = "TableStyleLight8"
)
### Write Summaries ###
#write Range summary stats
addWorksheet(data.xlsx, "Range Summary")
range_summary <- data.frame(
  Metric = c(
    "Range Extent Observed Area Min (km2)",
    "Range Extent Inferred Area Min (km2)",
    "Range Extent Inferred Area Max (km2)",
    "BGC Mapped Area Min (km2)",
    "BGC Mapped Area Min (percent)",
    "BGC Mapped Area Max (km2)",
    "BGC Mapped Area Max (%)",
    "BGC Range Min (km2)",
    "BGC Range Max (km2)"
  ),
  Value = c(
    round$BGC.MCP.Obs.Min.area.km2,
    round$BGC.MCP.Inf.Min.area.km2,
    round$BGC.MCP.Inf.Max.area.km2,
    round$BGC.Range.Mapped.Min.km2,
    round$BGC.Range.Mapped.Min.Percent,
    round$BGC.Range.Mapped.Max.km2,
    round$BGC.Range.Mapped.Max.Percent,
    round$BGC.Range.Min.km2,
    round$BGC.Range.Max.km2
  ))
writeData(data.xlsx, "Range Summary", range_summary, startRow = 1)
setColWidths(data.xlsx, sheet = "Range Summary", cols = 1:2, widths = "auto")
#write AOO summary stats
addWorksheet(data.xlsx, "AOO Summary")
aoo_summary <- data.frame(
  Metric = c(
    "Observed AOO Min (km2)",
    "Observed AOO Max (km2)",
    "Estimated AOO Min (km2)",
    "Estimated AOO Max (km2)",
    "Average Patch Size (ha)",
    "Maximum Patch Size (ha)",
    "Median Patch Size (ha)"
  ),
  Value = c(
    round$AOO.Obs.Min.km2,
    round$AOO.Obs.Max.km2,
    round$AOO.Est.Min.km2,
    round$AOO.Est.Max.km2,
    round$Patch.Size.Avg.ha,
    round$Patch.Size.Max.ha,
    round$Patch.Size.Med.ha
  ))
writeData(data.xlsx, "AOO Summary", aoo_summary, startRow = 1)
setColWidths(data.xlsx, sheet = "AOO Summary", cols = 1:2, widths = "auto")
#write NOO summary stats
addWorksheet(data.xlsx, "NOO Summary")
noo_summary <- data.frame(
  Parameter = c("NOO Minimum Count", 
                "NOO Maximum Count", 
                "Minimum Occurrence Size (km2)", 
                "Separation Distance (m)"),
  Value = c(round$NOO.Count.Min, 
            round$NOO.Count.Max, 
            round$min.occ.size.km2, 
            round$sep.dist.m))
writeData(data.xlsx, "NOO Summary", noo_summary, startRow = 1)
setColWidths(data.xlsx, sheet = "NOO Summary", cols = 1:2, widths = "auto")
# Write Impacts Summary
addWorksheet(data.xlsx, "Impact Summary")
writeData(data.xlsx, sheet = "Impact Summary", x = impact.summary.df)
setColWidths(data.xlsx, sheet = "Impact Summary", cols = 1:3, widths = "auto")
######################################################## FACTOR COMMENTS #######################################################
### assign text variables
Comments.Range.Extent <- glue("The observed range extent of this ecological community is {round$BGC.MCP.Obs.Min.area.km2} km2. It was calculated based on the area of a minimum convex polygon around all extant records of the community in ecosystem mapping (Ministry of Water, Land and Resource Stewardship n.d.) and confirmed ecosystem plot locations (BC Ministry of Forests n.d.). The observed range extent is an underestimate because inventory for this ecological community is incomplete. We infer that the range extent of this ecological community is between {round$BGC.MCP.Inf.Min.area.km2} km2 and {round$BGC.MCP.Inf.Max.area.km2} km2. The inferences are based on the area within minimum convex polygons around the boundaries of biogeoclimatic (BGC) units where the ecological community is known to occur (minimum), and around all BGC unit boundaries where it is known to occur plus those where it has been reported but not yet confirmed (maximum). The area within BGC units where the ecological community is known to occur is known as the BGC range. We estimate that the BGC range is between {round$BGC.Range.Min.km2} km2 and {round$BGC.Range.Max.km2} km2 based on the same subsets of BGC units used to calculate the minimum and maximum range extent above. The BGC range of this ecological community was determined based on expert observations (Deb MacKillop pers. com., Kristi Iverson pers. com., Kyla Rushton pers. com., and Harry Williams pers. com.), ecosystem plot locations from the BECMaster database (Ministry of Forests, n.d.), ecosystem mapping records (Ministry of Water, Land and Resource Stewardship n.d.), and from Biogeoclimatic Ecosystem Classification publications (MacKenzie and Moran 2004, MacKillop and Ehman 2016, MacKillop et al. 2018 and 2021, and Ryan et al. 2022). The BGC range mapping is based on BEC Version 13 (Ministry of Forests, n.d.).")
Comments.AOO <- glue("The observed area of occupancy (AOO) for this ecological community is between {round$AOO.Obs.Min.km2} km2 and {round$AOO.Obs.Max.km2} km2. The observed AOO values are based on the sums of the areas of the ecological community in ecosystem mapping records in BGC units where the ecological community is known to occur (minimum), and the sums of the areas of the ecological community in all BGC units where it is known to occur plus those where it has been reported but not yet confirmed. Five TEM Projects (BAPID numbers: 4, 216, 4523, 4917, 6536,6605) recorded occupancy rates for this ecological community across five BGC units (ESSFdc2, IDFdk2, SBSdk, SBSmc2, SBSmk1). The rate of occupancy ranged from 0.03% to 1.02% (excluding obvious outliers, or instances where the mapping was geographically constrained in a manner that would bias occupancy rates). The estimated AOO is between {round$AOO.Est.Min.km2} km2, and {round$AOO.Est.Max.km2} km2. The estimates were calculated by multiplying the area of the minimum and maximum estimated BGC range of the ecological community by minimum and maximum occupancy rates. It is challenging to determine how much inventory capable of recording this ecological community has occurred. Though a great deal of the BGC range is covered by ecosystem mapping project boundaries, most of it is predictive ecosystem mapping that does not identify wetland ecosystems to the site association level. Only {round$BGC.Range.Mapped.Min.Percent} percent of the BGC range is covered by mapping projects that recorded this ecological community and can therefore be confirmed as valid inventory. Up to {round$BGC.Range.Mapped.Max.Percent} percent of the BGC range is covered by other ecosystem mapping projects that appear capable of detecting the community.")
Comments.NOO <- glue("There are between {round$NOO.Count.Min} and {round$NOO.Count.Max} observed occurrences of this ecological community based on counts of clusters of extant records with a minimum patch size of 2 ha and a separation distance of {round$sep.dist.m} meters. The observed NOO values are based on counts of clusters of ecosystem mapping records and plot locations in BGC units where the community is known to occur (minimum), and records from all BGC units where it is known to occur plus those where it has been reported but not yet confirmed. Based on the high number of records and the low rate of inventory for this ecological community, we infer that there are greater than 300 occurrences total.")
Comments.NOOGEI <- glue("Insufficient data. Factor not assessed.")
Comments.AOOGEI <- glue("Insufficient data. Factor not assessed.")
Comments.Env.Spe <- glue("Factor not assessed.")
Comments.Assigned.Threats <- glue("Overall, {impact.summary.df$percent[impact.summary.df$CEF_DISTURB_GROUP == 'Cutblocks']}% of ecosystem mapping polygons (Ministry of Water, Land and Resource Stewardship n.d.) containing this ecological community intersect timber harvest polygons (Cumulative Effects Framework 2023). Though not the target of timber harvest, the ecological integrity of occurrences of this community may be impacted by ongoing timber harvest activity. The severity and scope of impacts from timber harvest are not known. Potential Impacts from climate change are unknown. This community has a large BGC range and climate change impacts may vary geographically, so the scope of climate change threat is unknown. Likewise, the severity of climate change impacts is unknown as we could not find information on how vulnerable this community is to various potential climate impacts (e.g., increased temperature, decreased precipitation, and changes in seasonality). It is possible that hydroclimatic conditions may support the expansion of the area suitable for this ecological community in parts of it’s biogeoclimatic range (Rodrigues et al. 2025). More geographic occurrence data is required to confidently determine trends and threats for this community.")
Comments.Calculated.Threats <- glue("")
Comments.Int.Vul <- glue("Factor not assessed.")
Comments.Short.Term.Trends <- glue("Insufficient data. Factor not assessed.")
Comments.Long.Term.Trends <- glue("Overall, 46% of ecosystem mapping polygons (Ministry of Water, Land and Resource Stewardship n.d.) containing this ecological community intersect timber harvest polygons (Cumulative Effects Framework 2023). Though not the target of timber harvest, the ecological integrity of occurrences of this community may be impacted by ongoing timber harvest activity. A Review of satellite imagery from 1984 to present suggests that known occurrences have been preserved. Obvious destruction or conversion from timber harvest were not observed. Some early cut blocks adjacent to occurrences have regenerated while new more recent cut blocks have been made adjacent to occurrences. Other human impacts such as mining, transmission lines, and residential development are smaller in scope (Cumulative Effects Framework 2023), but likely contribute to localized degradation of some occurrences of this ecological community. Increasing road density is a feature common to various industrial/residential human use in the last 200 years. The effects of roads on wetlands are varied and depend how and where they are constructed (Adamus 2014).We suspect that there has been a decline in the number of occurrences with good ecological integrity over the long-term period. This assumes that timber harvest and fire suppression have significantly altered the conditions of forest stands adjacent to this ecological community across it’s BGC range in the last 200 years. Timber harvest can decrease local evapotranspiration and lift local water tables, potentially increasing the total area of conditions suitable for the development of this ecological community (Adamus 2014). However, this effect may be temporary. Overall, it seems safe to suspect that conversion and altered hydrology from timber harvest and road building have degraded more occurrences (Adamus 2014) than would have been created, but supporting evidence is absent at this time.
")
#Vector of variable names and row numbers
factor.comments <- c(
  "Comments.Range.Extent",
  "Comments.AOO",
  "Comments.NOO",
  "Comments.NOOGEI",
  "Comments.AOOGEI",
  "Comments.Env.Spe",
  "Comments.Assigned.Threats",
  "Comments.Calculated.Threats",
  "Comments.Int.Vul",
  "Comments.Short.Term.Trends",
  "Comments.Long.Term.Trends"
)
factor.comments.rows <- c(1,3,6,9,10,11,12,13,14,15,16)

addWorksheet(data.xlsx, "Factor Comments")
# Clear  sheet to avoid ghost data
deleteData(data.xlsx, sheet = "Factor Comments", cols = 1:2, rows = 1:100, gridExpand = TRUE)

# Write explicitly row by row
for (i in seq_along(factor.comments)) {
  var_name <- factor.comments[i]
  var_value <- get(var_name)
  row_num <- factor.comments.rows[i]
  
  # Write the Label (Col 1)
  writeData(data.xlsx, sheet = "Factor Comments", x = var_name, startCol = 1, startRow = row_num)
  
  # Write the Comment (Col 2) - using 'as.character' to strip any weird formatting
  writeData(data.xlsx, sheet = "Factor Comments", x = as.character(var_value), startCol = 2, startRow = row_num)
}

# Apply wrap style
wrapStyle <- createStyle(wrapText = TRUE, valign = "top")
addStyle(data.xlsx, sheet = "Factor Comments", style = wrapStyle, 
         rows = factor.comments.rows, cols = 2, gridExpand = TRUE)

# Set fixed widths 
setColWidths(data.xlsx, sheet = "Factor Comments", cols = 1, widths = 30)
setColWidths(data.xlsx, sheet = "Factor Comments", cols = 2, widths = 200)

### Save workbook ###
saveWorkbook(
  data.xlsx,
  file.path(getwd(), "outputs", paste0(El.Sub.ID, "-data-csa-", format(Sys.time(), "%Y%m%d_%H%M"), ".xlsx")),
  overwrite = TRUE
)
#################################### To do next #####################################################################
# add commas and additional formatting to numbers for reporting under the 'round' list
# add QGIS styling
# Optimize presentation of impacts table by dropping/arranging columns
#### write functions and use input tables to set query values for
# 1. Biogeoclimatic range list
# 2. BAPIDS list
# 3. TEM Query and TEM AOO sum values