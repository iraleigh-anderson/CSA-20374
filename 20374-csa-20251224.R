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
BGC.Units.Min <- c("ESSFdc1", "ESSFdc2", "ESSFmh", "ESSFwm2", "ESSFxc1","ESSFxc2", "ICHdw4", "ICHmk1", "ICHmk2", "ICHmk3", "ICHmw3", "ICHmw5", "ICHvc", "ICHwk1", "ICHxm1", "IDFdk1", "IDFdk2", "IDFdk5", "IDFdm1", "IDFdm2", "IDFxk", "MSdk", "MSdm1", "MSdm2", "MSdm3", "MSdw", "MSxk1", "MSxk2", "SBPSdc", "SBPSmk", "SBPSxc","SBSdh1", "SBSdk", "SBSdw2", "SBSdw3", "SBSmc2", "SBSmk1", "SBSmw", "SBSwk1")
# minimal evidence supports occurrence of Ws04 in the BWBS. LMH 52 does not list Ws04 in the BWBS or SWB.  LMH68 does not list Ws04 or related Fl05 in any BWBS.  BWBSmw and wk1 seem to be at the NE margin of S. drummondiana and Spiraea douglasii range. Two plots support the occurrence of the Ws04 in the SBSdh1 (Kyla Rushton personal communication)
BGC.Units.Max <- c("BWBSwk1", "BWBSmw", "ESSFdc1", "ESSFdc2", "ESSFmh", "ESSFwm2", "ESSFxc1","ESSFxc2", "ICHdw4", "ICHmk1", "ICHmk2", "ICHmk3", "ICHmw3", "ICHmw5", "ICHvc", "ICHwk1", "ICHxm1", "IDFdk1", "IDFdk2", "IDFdk5", "IDFdm1", "IDFdm2", "IDFxk", "MSdk", "MSdm1", "MSdm2", "MSdm3", "MSdw", "MSxk1", "MSxk2", "SBPSdc", "SBPSmk", "SBPSxc","SBSdh1", "SBSdk", "SBSdw2", "SBSdw3", "SBSmc2", "SBSmk1", "SBSmw", "SBSwk1", "SBSwk2")

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
# It includes all TEI datasets that could conceivably detect the 20374, though it is uncertain whether all
# of these projects effectively inventory for the 20374. This only includes the core BGC Range, not the
# BGCs where occurrence of the 20374 is yet to be confirmed with confidence.
TEI.proj.bound.max <- st_read(
  dsn   = TEI.proj.bound.path,
  query = "
    SELECT *
    FROM WHSE_TERRESTRIAL_ECOLOGY_STE_TEI_PROJECT_BOUNDARIES_SP
    WHERE BUSINESS_AREA_PROJECT_ID IN (1, 2, 4, 9, 27, 28, 79, 91, 108, 114, 115, 119, 135, 145, 149, 183, 184, 192, 196, 198, 208, 209, 213, 214, 216, 218, 219, 221, 222, 226, 231, 233, 239, 240, 244, 247, 248, 1027, 1037, 1046, 1048, 1049, 1054, 1055, 1058, 1059, 1060, 1068, 1070, 4479, 4482, 4498, 4508, 4523, 4917, 5679, 5680, 6130, 6131, 6473, 6484, 6526, 6536, 6537, 6605, 6624)
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
    WHERE BUSINESS_AREA_PROJECT_ID IN (4, 108, 135, 209, 216, 233, 239, 244, 1046, 1048, 1049, 1054, 1055, 4523, 6473, 6536, 6605, 6624)
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
WHERE BAPID IN (4, 108, 135, 209, 216, 233, 239, 244, 1046, 1048, 1049, 1054, 1055)
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
vals <- c("DS", "WS", "WD")

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
  (TEM.m$BGC.Full == 'ESSFxc' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'ESSFdc2' & TEM.m$BAPID == 108 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 135 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'SBSdk' & TEM.m$BAPID == 209 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSdk' & TEM.m$BAPID == 216 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmc2' & TEM.m$BAPID == 216 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 233 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'IDFdm1' & TEM.m$BAPID == 239 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'MSdm1' & TEM.m$BAPID == 239 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'SBSmw' & TEM.m$BAPID == 244 & (TEM.m$SITEMC_S1 == 'WD' | TEM.m$SITEMC_S2 == 'WD' | TEM.m$SITEMC_S3 == 'WD')) |
  (TEM.m$BGC.Full == 'SBSwk1' & TEM.m$BAPID == 1046 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBSdw2' & TEM.m$BAPID == 1048 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBPSmk' & TEM.m$BAPID == 1049 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'SBPSxc' & TEM.m$BAPID == 1054 & (TEM.m$SITEMC_S1 == 'DS' | TEM.m$SITEMC_S2 == 'DS' | TEM.m$SITEMC_S3 == 'DS')) |
  (TEM.m$BGC.Full == 'MSdm1' & TEM.m$BAPID == 1055 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS')) |
  (TEM.m$BGC.Full == 'IDFdm1' & TEM.m$BAPID == 1055 & (TEM.m$SITEMC_S1 == 'WS' | TEM.m$SITEMC_S2 == 'WS' | TEM.m$SITEMC_S3 == 'WS'))

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
      BGC.Full == 'SBSdk' & BAPID %in% c(4, 209, 216) & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdk' & BAPID %in% c(4, 209, 216) & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'SBSdk' & BAPID %in% c(4, 209, 216) & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
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
      
      # ESSFdc2 / WS (BAPID 108)
      BGC.Full == 'ESSFdc2' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'ESSFdc2' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'ESSFdc2' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # ESSFxc / WS (BAPID 108)
      BGC.Full == 'ESSFxc' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'ESSFxc' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'ESSFxc' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # IDFdk2 / WS (BAPID 108)
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdk2' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # IDFdm1 / WS (BAPID 239, 1055)
      BGC.Full == 'IDFdm1' & BAPID %in% c(239, 1055) & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdm1' & BAPID %in% c(239, 1055) & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'IDFdm1' & BAPID %in% c(239, 1055) & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSdm1 / WS (BAPID 239, 1055)
      BGC.Full == 'MSdm1' & BAPID %in% c(239, 1055) & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm1' & BAPID %in% c(239, 1055) & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm1' & BAPID %in% c(239, 1055) & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSxk / WS (BAPID 108)
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'MSxk' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
      
      # MSdm2 / WS (BAPID 108)
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S1 == 'WS' ~ SDEC_1 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S2 == 'WS' ~ SDEC_2 * 0.1 * Shape_Area,
      BGC.Full == 'MSdm2' & BAPID == 108 & SITEMC_S3 == 'WS' ~ SDEC_3 * 0.1 * Shape_Area,
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
################################################  AOO Estimate and Inference ################################################
## Five TEM Projects (BAPID numbers: 4, 216, 4523, 4917, 6536,6605) recorded occupancy rates for this ecological community across five BGC units (ESSFdc2, IDFdk2, SBSdk, SBSmc2, SBSmk1). The rate of occupancy ranged from 0.03% to 1.02%.   THese rates of occupancy are multiplied by the min and max BGC range estimates to calculate an estimated area of occupancy below.
# BGC.Range.Mapped.Percent and BGC.Range.Mapped.km2 are defined above
AOO.Est.Min.km2 <- (BGC.Range.Min.km2*0.0003)
AOO.Est.Max.km2 <- (BGC.Range.Max.km2*0.0102)
#
#####################################################  NOO  #################################################
# Parameters
sep.dist.m <- 1000 #m
min.occ.size.km2 <- 0.0005 # 2 ha
### Create a new object of all occurrences merging features from BEC.Master.Focal.sf and TEM. Note that unvetted BEC plots are filtered at this stage, but furhter filtering of TEM occurs in the NOO process. 
# Filter to vetted plots
BEC.Vetted.Filtered <- BEC.Master.Focal.sf %>% 
  filter(Include == "Yes")
# combine with TEM
Occurrences <- bind_rows(BEC.Vetted.Filtered, TEM)
### minimum NOO estimate while excluding TEM from the BWBS. See comments in Range extent. ###
valid.occurrences.min <- TEM %>% filter(TEM$BGC_ZONE != 'BWBS' & TEM$BGC_ZONE != 'SBSwk2')
# Filter polygons for min size
min_occ <- valid.occurrences.min[valid.occurrences.min$AOO.km2 >= min.occ.size.km2, ]
# remove empty records
min_occ_clean <- min_occ[!st_is_empty(min_occ), ]
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
NOO.Count.Min <- aggregate(occ_buf, by = "GroupID", dissolve = TRUE)
# convert back to sf
NOO.Count.Min <- st_as_sf(NOO.Count.Min)
### maximum NOO estimate including the BWBS. See comments in Range extent. ###
valid.occurrences.max <- TEM

# Filter polygons for min size
max_occ <- valid.occurrences.max[valid.occurrences.max$AOO.km2 >= min.occ.size.km2, ]
# remove empty records
max_occ_clean <- max_occ[!st_is_empty(max_occ), ]
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
NOO.Count.Max <- aggregate(occ_buf, by = "GroupID", dissolve = TRUE)
# convert back to sf
NOO.Count.Max <- st_as_sf(NOO.Count.Max)

# Attach Group IDs to TEM
TEM <- TEM %>%
  st_join(NOO.Count.Min %>% select(GroupID_Min = GroupID))
TEM <- TEM %>%
  st_join(NOO.Count.Max %>% select(GroupID_Max = GroupID))

#### Count NOO
NOO.Min.Count <- nrow(NOO.Count.Min)
NOO.Max.Count <- nrow(NOO.Count.Max)
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
impact.summary <- Impacts %>%
  group_by(CEF_DISTURB_GROUP) %>%
  summarize(
    unique_teis_count = n_distinct(TEIS_ID, na.rm = FALSE)
  )%>%
  mutate(
    percent = (unique_teis_count / nrow(TEM)) * 100
  )

#################################################### WRITE SPATIAL #################################################################
# define output path
out.gpkg <- file.path(getwd(), "outputs", paste0(El.Sub.ID, "-csa-", format(Sys.time(), "%Y%m%d_%H%M"), ".gpkg"))

#write BEC Master points GPKG
st_write(BEC.Master.Focal.sf.buffer, out.gpkg, layer = "BECMaster")
#write TEM
st_write(TEM, out.gpkg, layer = "TEM")
# Write Range
st_write(BGC.Range.Min, out.gpkg, layer ="BGC.Range.Min")
st_write(BGC.Range.Max, out.gpkg, layer ="BGC.Range.max")
st_write(BGC.MCP.Inf.Min, out.gpkg, layer = "BGC.MCP.Inf.Min")
st_write(BGC.MCP.Inf.Max, out.gpkg, layer = "BGC.MCP.Inf.Max")
st_write(BGC.MCP.Obs.Min, out.gpkg, layer = "BGC.MCP.Obs.Min")

# Write NOO
st_write(NOO.Count.Max, out.gpkg, layer = "NOO.Max")
st_write(NOO.Count.Min, out.gpkg, layer = "NOO.Min")
# Write selected project boundaries and write ALL project boundaries (for ease of review)
st_write(TEI.proj.bound.all, out.gpkg, layer = "TEI.proj.bound.all")
st_write(TEI.proj.bound.min, out.gpkg, layer = "TEI.proj.bound.min")
st_write(TEI.proj.bound.max, out.gpkg, layer = "TEI.proj.bound.max")

############################################################### WRITE TABLES ####################################################
data.xlsx <- createWorkbook()
### WRITE DATA ###
### BEC Master Data 
addWorksheet(data.xlsx, "BEC.Master.Focal")
writeDataTable(data.xlsx, sheet = "BEC.Master.Focal", x = BEC.Master.Focal, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")
### Project Boundaries Data 
TEI.df <- st_drop_geometry(TEI.proj.bound.all)
addWorksheet(data.xlsx, "TEI Projects")
writeDataTable(data.xlsx, sheet = "TEI Projects", x = TEI.df, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")
### TEM Data 
TEM.df <- st_drop_geometry(TEM)
addWorksheet(data.xlsx, "TEM")
writeDataTable(data.xlsx, sheet = "TEM", x = TEM.df, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")
### Impacts Data
Impacts.df <- st_drop_geometry(Impacts)
addWorksheet(data.xlsx, "Impacts")
writeDataTable(data.xlsx, sheet = "Impacts", x = Impacts.df, 
               startCol = 1, startRow = 1, tableStyle = "TableStyleLight8")
# expand column widths
setColWidths(data.xlsx, sheet = "BEC.Master.Focal", cols = 1:50, widths = "auto")
setColWidths(data.xlsx, sheet = "TEI Projects", cols = 1:50, widths = "auto")
setColWidths(data.xlsx, sheet = "TEM", cols = 1:50, widths = "auto")
setColWidths(data.xlsx, sheet = "Impacts", cols = 1:50, widths = "auto")
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
addWorksheet(data.xlsx, "Geographic")

writeDataTable(
  wb         = data.xlsx,
  sheet      = "Geographic",
  x          = Geographic.Summary,
  startCol   = 1,
  startRow   = 1,
  tableStyle = "TableStyleLight8"
)
### Write Summaries ###
#write Range summary stats
addWorksheet(data.xlsx, "Range")
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
  Value = round(c(
    BGC.MCP.Obs.Min.area.km2,
    BGC.MCP.Inf.Min.area.km2,
    BGC.MCP.Inf.Max.area.km2,
    BGC.Range.Mapped.Min.km2,
    BGC.Range.Mapped.Min.Percent,
    BGC.Range.Mapped.Max.km2,
    BGC.Range.Mapped.Max.Percent,
    BGC.Range.Min.km2,
    BGC.Range.Max.km2
  ),1))
writeData(data.xlsx, "Range", range_summary, startRow = 1)
#write AOO summary stats
addWorksheet(data.xlsx, "AOO")
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
  Value = round(c(
    AOO.Obs.Min.km2,
    AOO.Obs.Max.km2,
    AOO.Est.Min.km2,
    AOO.Est.Max.km2,
    Patch.Size.Avg.ha,
    Patch.Size.Max.ha,
    Patch.Size.Med.ha
  ),1))
writeData(data.xlsx, "AOO", aoo_summary, startRow = 1)
#write NOO summary stats
addWorksheet(data.xlsx, "NOO")
noo_summary <- data.frame(
  Parameter = c("NOO Minimum Count", 
                "NOO Maximum Count", 
                "Minimum Occurrence Size (km2)", 
                "Separation Distance (m)"),
  Value = round(c(NOO.Min.Count, 
            NOO.Max.Count, 
            min.occ.size.km2, 
            sep.dist.m),1))
writeData(data.xlsx, "NOO", noo_summary, startRow = 1)
#Impacts Summary
impact.summary.df <- st_drop_geometry(impact.summary) %>%
  mutate(percent = round(percent, 1))
addWorksheet(data.xlsx, "Impact Summary")
writeData(data.xlsx, sheet = "Impact Summary", x = impact.summary.df)

# expand column widths
setColWidths(data.xlsx, sheet = "Range", cols = 1:2, widths = "auto")
setColWidths(data.xlsx, sheet = "AOO", cols = 1:2, widths = "auto")
setColWidths(data.xlsx, sheet = "NOO", cols = 1:2, widths = "auto")
setColWidths(data.xlsx, sheet = "Impact Summary", cols = 1:2, widths = "auto")

######################################################## FACTOR COMMENTS #######################################################
# assign text variables
Comments.Range.Extent <- glue("We estimate that the range extent of this ecological community is between {BGC.MCP.Inf.Min.area.km2} km2 and {BGC.MCP.Inf.Max.area.km2} km2. The estimates were made based on the area within minimum convex polygons around the boundaries of bigeoclimatic (BGC) units where the ecological community is known to occur (minimum), and around the all BGC unit boundaries where it is known to occur plus those where it has been reported but not yet confirmed (maximum). The area within BGC units where the ecological community is known to occur is known as the BGC range. We estimate that the BGC range is between {BGC.Range.Min} km2 and {BGC.Range.Max} km2 based on the same subsets of BGC units used to calculate the minimum and maximum range extent above. The BGC range of this ecological community was determined based on expert observations (Deb MacKillop pers. com., Kristi Iverson pers. com., Kyla Rushton pers. com., and Harry Williams pers. com.), ecosystem plot locations from the BECMaster database (Ministry of Forests, n.d.), ecosystem mapping records (Ministry of Water, Land and Resource Stewardship n.d.), and from Biogeoclimatic Ecosystem Classification Publications (MacKenzie and Moran 2004, MacKillop and Ehman 2016, MacKillop et al. 2018 and 2021, and Ryan et al. 2022). The mapping is based on BEC Version 12 (Ministry of Forests, n.d.). ")
Comments.AOO <- glue("The observed area of occupancy (AOO) for this ecological community is between {AOO.Obs.Min.km2} km2 and {AOO.Obs.Max.km2} km2.  The observed AOO values are based on the sums of the areas of the ecological community in ecosystem mapping records in BGC units where the ecological community is known to occur (minimum), and in all BGC units where it is known to occur plus those where it has been reported but not yet confirmed. The estimated AOO is between {AOO.Est.Min.km2} km2, and {AOO.Est.Max.km2} km2. The estimates were calculated by multiplying the area of of the estimated BGC range of the ecological community by minimum and maximum occupancy rates of the ecological community recorded in ecosystem mappping projects (i.e., area of the ecological community mapped per total area mapped). It is challenging to determine how much inventory capable of recording this ecological community has occurred. Though a great deal of the biogeoclimatic range has ecosystem mapping, most of it is predictive ecosystem mapping that does not identify wetland ecosystems to the site association level. Only {BGC.Range.Mapped.Min.Percent} percent of the BGC range is covered by mapping projects that actually recorded this ecological community and can therefore be confirmed as valid inventory. Up to {BGC.Range.Mapped.Max.Percent} percent of the BGC range of the ecological community is covered by terrestrial ecosystem mapping projects that appear capable of mapping the focal ecological community but did not actually record it in any polygons. It appears that effective inventory for this ecological community is minimal.").
Comments.NOO <- glue("There are between {NOO.Min.Count} and {NOO.Max.Count} observed occurrences of this ecological community based on counts of clusters of ecosystem mapping records with a minimum patch size of {min.occ.size.km2} km2 and a separation distance of {sep.dist.m} meters. The observed NOO values are based on running the clustering algorithm based on ecosystem mapping records in BGC units where the ecological community is known to occur (minimum), and records from all BGC units where it is known to occur plus those where it has been reported but not yet confirmed.  Based on the apparent low rate of inventory for this ecological community, we infer that there may be greater than 300 occurrences total.")
Comments.NOOGEI <- glue("Insufficient data. Factor not assessed.")
Comments.AOOGEI <- glue("Insufficient data. Factor not assessed.")
Comments.Env.Spe <- glue("Factor not assessed.")
Comments.Assigned.Threats <- glue("Fifty-one percent of ecosystem mapping polygons (Ministry of Water, Land and Resource Stewardship n.d.) containing this ecological community intersect timber harvest cutblock polygons (Ministry of Environment and Climate Change Strategy 2023). Though not the target of timber timber harvest, the ecological integrity of occurrences of this community may be impacted by ongoing timber harvest activity. The severity of potential impacts from timber harvest are not known.")
Comments.Calculated.Threats <- glue("")
Comments.Int.Vul <- glue("Factor not assessed.")
Comments.Short.Term.Trends <- glue("Insufficient data. Factor not assessed.")
Comments.Long.Term.Trends <- glue("Insufficient data. Factor not assessed.")
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

addWorksheet(data.xlsx, "Factor.Comments")

# Write variable names and values to specific rows to allow single paste into the rank calculator
for (i in seq_along(factor.comments)) {
  var_name <- factor.comments[i]
  var_value <- get(var_name, envir = .GlobalEnv)
  row_num <- factor.comments.rows[i]
  
  writeData(data.xlsx, sheet = "Factor.Comments", x = var_name,  startCol = 1, startRow = row_num)
  writeData(data.xlsx, sheet = "Factor.Comments", x = var_value, startCol = 2, startRow = row_num)
}
#auto expand column widths and apply text wrapping
wrapStyle <- createStyle(wrapText = TRUE)
addStyle(data.xlsx, sheet = "Factor.Comments", style = wrapStyle, rows = 1:100, cols = 2:2, gridExpand = TRUE)
setColWidths(data.xlsx, sheet = "Factor.Comments", cols = 1:50, widths = "auto")

### Save workbook ###
saveWorkbook(
  data.xlsx,
  file.path(getwd(), "outputs", paste0(El.Sub.ID, "-data-csa-", format(Sys.time(), "%Y%m%d_%H%M"), ".xlsx")),
  overwrite = TRUE
)

#################################### To do next #####################################################################
# Get BGC range observed.
# Get NOO with BECmaster included. Double check filter TERMS re SBSwk2 in ZONE field (error)
# clean up AOO calculation based on my spreadhseet
# update factor comments
# fix variable names
# add rounding
# add NULL impact label to impact summary table
