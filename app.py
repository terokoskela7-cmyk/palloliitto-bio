"""Palloliitto Bio: biologisen kypsyyden laskenta ja yli-ikäisyyssäännön tarkistus.

Laskenta noudattaa Palloliiton Excel-työkaluja:
  * "Mirwald ja Khamis-Roche Yli-ikäisyyssääntö pojat ja tytöt" (Mirwald-kaavat,
    kuukausikohtaiset raja-arvot ja säännön esimerkit)
  * "Khamis_Roche_NEW_1" (Khamis-Roche-kertoimet ja ruotsalainen viiteaineisto)
"""

import base64
import datetime
import json
import os
import math
import re
import traceback
from decimal import ROUND_HALF_UP, Decimal

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

SHEET_NAME = "Palloliitto Bio Tulokset"

# Sarakkeet, jotka sovellus tallentaa Google Sheetsiin (taulukon rivin 1 otsikot).
SHEET_HEADERS = [
    "Pvm",
    "Seura",
    "Testaaja",
    "Testaajan sähköposti",
    "Metodi",
    "PalloID",
    "Sukupuoli",
    "Syntymäkuukausi",
    "Ikäryhmä (K-R)",
    "Ikä",
    "Pituus ka (cm)",
    "Paino ka (kg)",
    "Istumapituus ka (cm)",
    "Penkin korkeus (cm)",
    "PHV-ikä / Bio-ikä",
    "Raja-arvo",
    "Pelioikeus",
    "Mittaus voimassa asti",
]

# Palloliiton ohje: biologisen iän mittaustulos on voimassa 6 kuukautta mittauspäivästä.
VALIDITY_MONTHS = 6
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# Penkin korkeus, jonka Excel vähentää istumapituusmittauksesta (Mirwald).
DEFAULT_BENCH_HEIGHT_CM = 40.0

# Mirwald, pojat: Excel ("Mirwald PHV ikä" -välilehti) käyttää vakiota -9.3236.
# Mirwald ym. (2002) -alkuperäisjulkaisussa vakio on -9.236. Sovellus noudattaa
# Palloliiton Exceliä, jotta tulokset täsmäävät. Muuta tarvittaessa tästä.
MIRWALD_BOYS_INTERCEPT = -9.3236
MIRWALD_GIRLS_INTERCEPT = -9.376

MONTHS_FI = [
    "Tammikuu",
    "Helmikuu",
    "Maaliskuu",
    "Huhtikuu",
    "Toukokuu",
    "Kesäkuu",
    "Heinäkuu",
    "Elokuu",
    "Syyskuu",
    "Lokakuu",
    "Marraskuu",
    "Joulukuu",
]

# Mirwald: PHV-iän tulee YLITTÄÄ raja-arvo (syntymäkuukauden mukaan).
MIRWALD_LIMITS = {
    "Poika": [
        14.966667,
        14.883333,
        14.800000,
        14.716667,
        14.633333,
        14.550000,
        14.466667,
        14.383333,
        14.300000,
        14.216667,
        14.133333,
        14.050000,
    ],
    "Tyttö": [
        13.066667,
        12.983333,
        12.900000,
        12.816667,
        12.733333,
        12.650000,
        12.566667,
        12.483333,
        12.400000,
        12.316667,
        12.233333,
        12.150000,
    ],
}

# Khamis-Roche: biologisen iän tulee ALITTAA raja-arvo (ikäryhmä + syntymäkuukausi).
# Arvot ovat "keskimääräinen kronologinen ikä vähennettynä 0,75 vuodella".
KR_LIMITS = {
    13: [11.75, 11.84, 11.93, 12.02, 12.11, 12.20, 12.30, 12.39, 12.48, 12.57, 12.66, 12.75],
    14: [12.75, 12.84, 12.93, 13.02, 13.11, 13.20, 13.30, 13.39, 13.48, 13.57, 13.66, 13.75],
    15: [13.75, 13.84, 13.93, 14.02, 14.11, 14.20, 14.30, 14.39, 14.48, 14.57, 14.66, 14.75],
}

# ---------------------------------------------------------------------------
# Khamis-Roche-kertoimet (Khamis_Roche_NEW_1.xlsx, välilehdet C_male ja C_female).
# Rivi: (ikä, intercept [in], pituuskerroin, painokerroin [lb], vanhempien pituus).
# Excel muuntaa: intercept*2.54 (cm), painokerroin*5.599741 (kg).
_KR_COEFFS_BOYS_RAW = [
    (4, -10.2567, 1.23812, -0.087235, 0.50286),
    (4.5, -10.719, 1.15964, -0.074454, 0.52887),
    (5, -11.0213, 1.10674, -0.064778, 0.53919),
    (5.5, -11.1556, 1.0748, -0.05776, 0.53691),
    (6, -11.1138, 1.05923, -0.052947, 0.52513),
    (6.5, -11.0221, 1.05542, -0.049892, 0.50692),
    (7, -10.9984, 1.05877, -0.048144, 0.48538),
    (7.5, -11.0214, 1.06467, -0.047256, 0.46361),
    (8, -11.0696, 1.06853, -0.046778, 0.44469),
    (8.5, -11.122, 1.06572, -0.046261, 0.43171),
    (9, -11.1571, 1.05166, -0.045254, 0.42776),
    (9.5, -11.1405, 1.02174, -0.043311, 0.43593),
    (10, -11.038, 0.97135, -0.039981, 0.45932),
    (10.5, -10.8286, 0.89589, -0.034814, 0.50101),
    (11, -10.4917, 0.81239, -0.02905, 0.54781),
    (11.5, -10.0065, 0.74134, -0.024167, 0.58409),
    (12, -9.3522, 0.68325, -0.020076, 0.60927),
    (12.5, -8.6055, 0.63869, -0.016681, 0.62279),
    (13, -7.8632, 0.60818, -0.013895, 0.62407),
    (13.5, -7.1348, 0.59228, -0.011624, 0.61253),
    (14, -6.4299, 0.59151, -0.009776, 0.58762),
    (14.5, -5.7578, 0.60643, -0.008261, 0.54875),
    (15, -5.1282, 0.63757, -0.006988, 0.49536),
    (15.5, -4.5092, 0.68548, -0.005863, 0.42687),
    (16, -3.9292, 0.75069, -0.004795, 0.34271),
    (16.5, -3.4873, 0.83375, -0.003695, 0.24231),
    (17, -3.283, 0.9352, -0.00247, 0.1251),
    (17.5, -3.4156, 1.05558, -0.001027, -0.0095),
]
_KR_COEFFS_GIRLS_RAW = [
    (4, -8.1325, 1.24768, -0.19435, 0.44774),
    (4.5, -6.47656, 1.22177, -0.18519, 0.41381),
    (5, -5.13583, 1.19932, -0.1753, 0.38467),
    (5.5, -4.13791, 1.1788, -0.16484, 0.36039),
    (6, -3.51039, 1.15866, -0.154, 0.34105),
    (6.5, -3.14322, 1.13737, -0.14294, 0.32672),
    (7, -2.87645, 1.11342, -0.13184, 0.31748),
    (7.5, -2.66291, 1.08525, -0.12086, 0.3134),
    (8, -2.45559, 1.05135, -0.11019, 0.31457),
    (8.5, -2.20728, 1.01018, -0.09999, 0.32105),
    (9, -1.87098, 0.9602, -0.09044, 0.33291),
    (9.5, -1.0633, 0.89989, -0.08171, 0.35025),
    (10, 0.33468, 0.82771, -0.07397, 0.37312),
    (10.5, 1.97366, 0.74213, -0.06739, 0.40161),
    (11, 3.50436, 0.67173, -0.06136, 0.42042),
    (11.5, 4.57747, 0.6415, -0.05518, 0.41686),
    (12, 4.84365, 0.64452, -0.04894, 0.3949),
    (12.5, 4.27869, 0.67386, -0.04272, 0.3585),
    (13, 3.21417, 0.7226, -0.03661, 0.31163),
    (13.5, 1.83456, 0.78383, -0.03067, 0.25826),
    (14, 0.32425, 0.85062, -0.025, 0.20235),
    (14.5, -1.13224, 0.91605, -0.01967, 0.14787),
    (15, -2.35055, 0.97319, -0.01477, 0.0988),
    (15.5, -3.10326, 1.01514, -0.01037, 0.05909),
    (16, -3.17885, 1.03496, -0.00655, 0.03272),
    (16.5, -2.41657, 1.02573, -0.0034, 0.02364),
    (17, -0.65579, 0.98054, -0.001, 0.03584),
    (17.5, 2.26429, 0.89246, 0.00057, 0.07327),
]
# Ruotsalainen viiteaineisto (Albertsson ym., välilehti SWE_REF): ikä, keskimääräinen %PAH, SD.
_SWE_REF_BOYS_RAW = [
    (0, 0.2814942082802195, 0.01158343956104857),
    (0.1, 0.30517098043562596, 0.012403702266807093),
    (0.2, 0.32884775259103255, 0.01322396497256556),
    (0.25, 0.34068613866873576, 0.013634096325444767),
    (0.3, 0.347846810397384, 0.01357867316965028),
    (0.4, 0.3621681538546805, 0.01346782685806136),
    (0.5, 0.37648949731197695, 0.01335698054647233),
    (0.6, 0.3866651887158455, 0.013423488333425715),
    (0.7, 0.396840880119714, 0.0134899961203791),
    (0.75, 0.40192872582164824, 0.013523250013855792),
    (0.8, 0.4059746161946461, 0.013678434850080412),
    (0.9, 0.41406639694064173, 0.013988804522529541),
    (1, 0.4221581776866375, 0.014299174194978614),
    (1.1, 0.4297289807681649, 0.01458737460511006),
    (1.2, 0.43729978384969237, 0.014875575015241393),
    (1.3, 0.44487058693121984, 0.015163775425372783),
    (1.4, 0.4524413900127473, 0.015451975835504062),
    (1.5, 0.46001219309427477, 0.01574017624563545),
    (1.6, 0.46575403203458404, 0.015972953499972298),
    (1.7, 0.4714958709748933, 0.0162057307543092),
    (1.8, 0.4772377099152026, 0.01643850800864599),
    (1.9, 0.48297954885551186, 0.016671285262982838),
    (2, 0.48872138779582114, 0.01690406251731974),
    (2.1, 0.4936207947680541, 0.017303109239040104),
    (2.2, 0.4985202017402871, 0.017702155960760413),
    (2.3, 0.5034196087125201, 0.01810120268248072),
    (2.4, 0.5083190156847531, 0.01850024940420103),
    (2.5, 0.513218422656986, 0.01889929612592145),
    (2.6, 0.518117829629219, 0.01929834284764176),
    (2.7, 0.523017236601452, 0.019697389569362067),
    (2.8, 0.5279166435736851, 0.020096436291082376),
    (2.9, 0.5328160505459181, 0.020495483012802684),
    (3, 0.5377154575181511, 0.020894529734522993),
    (3.1, 0.5421271407193926, 0.020999833730532602),
    (3.2, 0.546538823920634, 0.0211051377265421),
    (3.3, 0.5509505071218755, 0.0212104417225516),
    (3.4, 0.555362190323117, 0.021315745718561208),
    (3.5, 0.5597738735243585, 0.021421049714570706),
    (3.6, 0.5641855567256, 0.021526353710580204),
    (3.7, 0.5685972399268414, 0.021631657706589813),
    (3.8, 0.5730089231280829, 0.02173696170259931),
    (3.9, 0.5774206063293243, 0.02184226569860892),
    (4, 0.5818322895305659, 0.02194756969461842),
    (4.1, 0.585251898243086, 0.02235770104749768),
    (4.2, 0.588671506955606, 0.02276783240037694),
    (4.3, 0.5920911156681261, 0.023177963753256203),
    (4.4, 0.5955107243806462, 0.023588095106135354),
    (4.5, 0.5989303330931663, 0.023998226459014504),
    (4.6, 0.6023499418056864, 0.024408357811893877),
    (4.7, 0.6057695505182065, 0.024818489164773028),
    (4.8, 0.6091891592307266, 0.02522862051765218),
    (4.9, 0.6126087679432467, 0.02563875187053155),
    (5, 0.6160283766557668, 0.0260488832234107),
    (5.1, 0.6196364240979881, 0.026265033531009285),
    (5.2, 0.6232444715402095, 0.026481183838607758),
    (5.3, 0.6268525189824309, 0.026697334146206342),
    (5.4, 0.6304605664246522, 0.026913484453804815),
    (5.5, 0.6340686138668735, 0.027129634761403287),
    (5.6, 0.637676661309095, 0.02734578506900176),
    (5.7, 0.6412847087513163, 0.027561935376600344),
    (5.8, 0.6448927561935377, 0.027778085684198817),
    (5.9, 0.648500803635759, 0.0279942359917974),
    (6, 0.6521088510779803, 0.028210386299395873),
    (6.1, 0.6560106412459125, 0.028177132405919236),
    (6.2, 0.6599124314138447, 0.02814387851244249),
    (6.3, 0.6638142215817768, 0.02811062461896574),
    (6.4, 0.667716011749709, 0.028077370725489215),
    (6.5, 0.6716178019176412, 0.028044116832012467),
    (6.6, 0.6755195920855733, 0.02801086293853572),
    (6.7, 0.6794213822535056, 0.02797760904505897),
    (6.8, 0.6833231724214377, 0.027944355151582334),
    (6.9, 0.6872249625893698, 0.027911101258105697),
    (7, 0.691126752757302, 0.02787784736462895),
    (7.1, 0.694507565260766, 0.02824364019287262),
    (7.2, 0.6978883777642299, 0.028609433021116182),
    (7.3, 0.7012691902676939, 0.028975225849359854),
    (7.4, 0.7046500027711579, 0.029341018677603525),
    (7.5, 0.7080308152746217, 0.029706811505847197),
    (7.6, 0.7114116277780858, 0.030072604334090647),
    (7.7, 0.7147924402815496, 0.03043839716233432),
    (7.8, 0.7181732527850136, 0.03080418999057799),
    (7.9, 0.7215540652884777, 0.031169982818821662),
    (8, 0.7249348777919415, 0.031535775647065334),
    (8.1, 0.7275231391675442, 0.03167433353655158),
    (8.2, 0.7301114005431469, 0.031812891426037826),
    (8.3, 0.7326996619187498, 0.03195144931552385),
    (8.4, 0.7352879232943524, 0.03209000720501021),
    (8.5, 0.7378761846699551, 0.03222856509449645),
    (8.6, 0.7404644460455578, 0.0323671229839827),
    (8.7, 0.7430527074211605, 0.032505680873468945),
    (8.8, 0.7456409687967632, 0.03264423876295508),
    (8.9, 0.7482292301723661, 0.032782796652441326),
    (9, 0.7508174915479687, 0.03292135454192757),
    (9.1, 0.7542537272072272, 0.03327606273901229),
    (9.2, 0.7576899628664855, 0.03363077093609712),
    (9.3, 0.7611261985257441, 0.03398547913318184),
    (9.4, 0.7645624341850024, 0.03434018733026667),
    (9.5, 0.7679986698442609, 0.03469489552735128),
    (9.6, 0.7714349055035193, 0.03504960372443611),
    (9.7, 0.7748711411627777, 0.035404311921520826),
    (9.8, 0.7783073768220362, 0.035759020118605545),
    (9.9, 0.7817436124812945, 0.03611372831569026),
    (10, 0.785179848140553, 0.036468436512775204),
    (10.1, 0.7872748434295848, 0.03662362134899955),
    (10.2, 0.7893698387186165, 0.03677880618522411),
    (10.3, 0.7914648340076483, 0.03693399102144879),
    (10.4, 0.79355982929668, 0.03708917585767346),
    (10.5, 0.7956548245857119, 0.037244360693897915),
    (10.6, 0.7977498198747436, 0.03739954553012237),
    (10.7, 0.7998448151637753, 0.037554730366347155),
    (10.8, 0.801939810452807, 0.03770991520257172),
    (10.9, 0.8040348057418389, 0.03786510003879617),
    (11, 0.8061298010308706, 0.03802028487502085),
    (11.1, 0.8100260488832234, 0.038408246965582205),
    (11.2, 0.813922296735576, 0.03879620905614367),
    (11.3, 0.8178185445879288, 0.03918417114670503),
    (11.4, 0.8217147924402816, 0.03957213323726638),
    (11.5, 0.8256110402926341, 0.03996009532782796),
    (11.6, 0.8295072881449869, 0.040348057418389316),
    (11.7, 0.8334035359973395, 0.040736019508950894),
    (11.8, 0.8372997838496923, 0.04112398159951225),
    (11.9, 0.841196031702045, 0.041511943690073716),
    (12, 0.8450922795543977, 0.04189990578063518),
    (12.1, 0.8480851299673002, 0.04228232555561706),
    (12.2, 0.8510779803802027, 0.04266474533059916),
    (12.3, 0.8540708307931053, 0.04304716510558104),
    (12.4, 0.8570636812060077, 0.043429584880563143),
    (12.5, 0.8600565316189104, 0.043812004655545134),
    (12.6, 0.8630493820318128, 0.04419442443052701),
    (12.7, 0.8660422324447153, 0.044576844205509114),
    (12.8, 0.869035082857618, 0.04495926398049099),
    (12.9, 0.8720279332705204, 0.045341683755473094),
    (13, 0.8750207836834228, 0.045724103530455085),
    (13.1, 0.8799866984426092, 0.04585157678878227),
    (13.2, 0.8849526132017956, 0.04597905004710967),
    (13.3, 0.8899185279609821, 0.046106523305436964),
    (13.4, 0.8948844427201684, 0.04623399656376448),
    (13.5, 0.8998503574793549, 0.04636146982209166),
    (13.6, 0.9048162722385412, 0.046488943080419065),
    (13.7, 0.9097821869977275, 0.04661641633874636),
    (13.8, 0.914748101756914, 0.04674388959707365),
    (13.9, 0.9197140165161005, 0.046871362855400944),
    (14, 0.9246799312752868, 0.04699883611372824),
    (14.1, 0.9277836279997783, 0.046887989802139374),
    (14.2, 0.9308873247242698, 0.04677714349055029),
    (14.3, 0.9339910214487613, 0.04666629717896131),
    (14.4, 0.9370947181732528, 0.04655545086737234),
    (14.5, 0.9401984148977441, 0.046444604555783364),
    (14.6, 0.9433021116222357, 0.0463337582441945),
    (14.7, 0.9464058083467272, 0.046222911932605526),
    (14.8, 0.9495095050712187, 0.04611206562101644),
    (14.9, 0.9526132017957102, 0.04600121930942758),
    (15, 0.9557168985202017, 0.04589037299783849),
    (15.1, 0.958288532949066, 0.04541927617358543),
    (15.2, 0.9608601673779305, 0.04494817934933226),
    (15.3, 0.9634318018067948, 0.04447708252507898),
    (15.4, 0.9660034362356592, 0.04400598570082581),
    (15.5, 0.9685750706645235, 0.04353488887657253),
    (15.6, 0.9711467050933881, 0.04306379205231947),
    (15.7, 0.9737183395222525, 0.0425926952280663),
    (15.8, 0.9762899739511168, 0.04212159840381302),
    (15.9, 0.9788616083799812, 0.04165050157955996),
    (16, 0.9814332428088456, 0.04117940475530679),
    (16.1, 0.9823421825638752, 0.0408579504516986),
    (16.2, 0.9832511223189049, 0.04053649614809074),
    (16.3, 0.9841600620739345, 0.040215041844482546),
    (16.4, 0.9850690018289642, 0.039893587540874687),
    (16.5, 0.9859779415839938, 0.039572133237266494),
    (16.6, 0.9868868813390234, 0.03925067893365852),
    (16.7, 0.9877958210940531, 0.03892922463005044),
    (16.8, 0.9887047608490827, 0.03860777032644247),
    (16.9, 0.9896137006041125, 0.03828631602283428),
    (17, 0.9905226403591421, 0.037964861719226306),
    (17.1, 0.9914703763232278, 0.03781521919858122),
    (17.2, 0.9924181122873137, 0.03766557667793591),
    (17.3, 0.9933658482513994, 0.03751593415729104),
    (17.4, 0.9943135842154852, 0.03736629163664562),
    (17.5, 0.9952613201795709, 0.037216649116000755),
    (17.6, 0.9962090561436568, 0.037067006595355445),
    (17.7, 0.9971567921077426, 0.03691736407471036),
    (17.8, 0.9981045280718284, 0.03676772155406516),
    (17.9, 0.9990522640359143, 0.03661807903342007),
    (18, 1, 0.036468436512775204),
]
_SWE_REF_GIRLS_RAW = [
    (0, 0.30053398932021363, 0.012539749205015938),
    (0.1, 0.32474950500989985, 0.012899742005159864),
    (0.2, 0.348965020699586, 0.013259734805303902),
    (0.25, 0.36107277854442915, 0.013439731205375893),
    (0.3, 0.36859662806743865, 0.013475730485390303),
    (0.4, 0.38364432711345775, 0.013547729045419121),
    (0.5, 0.39869202615947685, 0.013619727605447884),
    (0.6, 0.4097318053638928, 0.013979720405591867),
    (0.7, 0.42077158456830865, 0.014339713205735849),
    (0.75, 0.4262914741705166, 0.014519709605807896),
    (0.8, 0.43064738705225897, 0.014723705525889419),
    (0.9, 0.43935921281574375, 0.015131697366052743),
    (1, 0.4480710385792285, 0.015539689206215901),
    (1.1, 0.45650686986260275, 0.01582768344633112),
    (1.2, 0.46494270114597713, 0.016115677686446228),
    (1.3, 0.4733785324293514, 0.016403671926561447),
    (1.4, 0.4818143637127258, 0.016691666166676666),
    (1.5, 0.49025019499610006, 0.01697966040679183),
    (1.6, 0.49676606467870643, 0.017387652246955043),
    (1.7, 0.5032819343613127, 0.0177956440871182),
    (1.8, 0.5097978040439192, 0.01820363592728147),
    (1.9, 0.5163136737265255, 0.01861162776744474),
    (2, 0.5228295434091319, 0.019019619607607785),
    (2.1, 0.5282114357712846, 0.01928961420771591),
    (2.2, 0.5335933281334374, 0.019559608807823814),
    (2.3, 0.5389752204955901, 0.01982960340793183),
    (2.4, 0.5443571128577428, 0.020099598008039843),
    (2.5, 0.5497390052198956, 0.020369592608147857),
    (2.6, 0.5551208975820484, 0.020639587208255872),
    (2.7, 0.5605027899442011, 0.020909581808363886),
    (2.8, 0.5658846823063539, 0.0211795764084719),
    (2.9, 0.5712665746685066, 0.021449571008579915),
    (3, 0.5766484670306594, 0.02171956560868782),
    (3.1, 0.5813703725925482, 0.02200155996880071),
    (3.2, 0.586092278154437, 0.02228355432891338),
    (3.3, 0.5908141837163258, 0.02256554868902616),
    (3.4, 0.5955360892782144, 0.022847543049139052),
    (3.5, 0.6002579948401032, 0.023129537409251832),
    (3.6, 0.604979900401992, 0.023411531769364613),
    (3.7, 0.6097018059638808, 0.023693526129477394),
    (3.8, 0.6144237115257696, 0.023975520489590174),
    (3.9, 0.6191456170876583, 0.024257514849703066),
    (4, 0.623867522649547, 0.024539509209815846),
    (4.1, 0.6279234415311694, 0.02488150236995268),
    (4.2, 0.6319793604127918, 0.0252234955300894),
    (4.3, 0.6360352792944142, 0.025565488690226124),
    (4.4, 0.6400911981760365, 0.025907481850362957),
    (4.5, 0.6441471170576589, 0.02624947501049979),
    (4.6, 0.6482030359392813, 0.026591468170636623),
    (4.7, 0.6522589548209037, 0.026933461330773345),
    (4.8, 0.656314873702526, 0.027275454490910178),
    (4.9, 0.6603707925841483, 0.0276174476510469),
    (5, 0.6644267114657707, 0.027959440811183733),
    (5.1, 0.6683866322673546, 0.028133437331253508),
    (5.2, 0.6723465530689386, 0.02830743385132306),
    (5.3, 0.6763064738705227, 0.028481430371392502),
    (5.4, 0.6802663946721066, 0.028655426891462166),
    (5.5, 0.6842263154736905, 0.02882942341153183),
    (5.6, 0.6881862362752745, 0.029003419931601382),
    (5.7, 0.6921461570768586, 0.029177416451670934),
    (5.8, 0.6961060778784425, 0.029351412971740487),
    (5.9, 0.7000659986800265, 0.02952540949181015),
    (6, 0.7040259194816104, 0.029699406011879814),
    (6.1, 0.7079858402831943, 0.029717405651887074),
    (6.2, 0.7119457610847784, 0.029735405291894113),
    (6.3, 0.7159056818863624, 0.029753404931901373),
    (6.4, 0.7198656026879463, 0.029771404571908633),
    (6.5, 0.7238255234895302, 0.029789404211915782),
    (6.6, 0.7277854442911142, 0.02980740385192293),
    (6.7, 0.7317453650926983, 0.02982540349193019),
    (6.8, 0.7357052858942822, 0.02984340313193734),
    (6.9, 0.7396652066958661, 0.0298614027719446),
    (7, 0.7436251274974501, 0.02987940241195164),
    (7.1, 0.7473630527389452, 0.03034739305213896),
    (7.2, 0.7511009779804404, 0.03081538369232617),
    (7.3, 0.7548389032219356, 0.03128337433251338),
    (7.4, 0.7585768284634308, 0.03175136497270048),
    (7.5, 0.762314753704926, 0.03221935561288769),
    (7.6, 0.766052678946421, 0.0326873462530749),
    (7.7, 0.7697906041879162, 0.03315533689326222),
    (7.8, 0.7735285294294113, 0.03362332753344943),
    (7.9, 0.7772664546709066, 0.03409131817363653),
    (8, 0.7810043799124017, 0.03455930881382374),
    (8.1, 0.7838363232735346, 0.03464930701385971),
    (8.2, 0.7866682666346673, 0.03473930521389568),
    (8.3, 0.7895002099958001, 0.034829303413931756),
    (8.4, 0.7923321533569329, 0.034919301613967724),
    (8.5, 0.7951640967180655, 0.0350092998140038),
    (8.6, 0.7979960400791983, 0.03509929801403977),
    (8.7, 0.8008279834403312, 0.03518929621407574),
    (8.8, 0.803659926801464, 0.035279294414111706),
    (8.9, 0.8064918701625967, 0.035369292614147785),
    (9, 0.8093238135237295, 0.03545929081418364),
    (9.1, 0.8132837343253134, 0.035855282894342144),
    (9.2, 0.8172436551268975, 0.036251274974500425),
    (9.3, 0.8212035759284815, 0.03664726705465893),
    (9.4, 0.8251634967300655, 0.03704325913481732),
    (9.5, 0.8291234175316494, 0.03743925121497571),
    (9.6, 0.8330833383332333, 0.03783524329513421),
    (9.7, 0.8370432591348174, 0.038231235375292494),
    (9.8, 0.8410031799364014, 0.038627227455450885),
    (9.9, 0.8449631007379854, 0.039023219535609166),
    (10, 0.8489230215395693, 0.03941921161576767),
    (10.1, 0.85172496550069, 0.039563208735825306),
    (10.2, 0.8545269094618109, 0.03970720585588283),
    (10.3, 0.8573288534229317, 0.03985120297594047),
    (10.4, 0.8601307973840524, 0.03999520009599811),
    (10.5, 0.8629327413451731, 0.040139197216055744),
    (10.6, 0.8657346853062939, 0.04028319433611327),
    (10.7, 0.8685366292674146, 0.04042719145617091),
    (10.8, 0.8713385732285355, 0.040571188576228545),
    (10.9, 0.8741405171896564, 0.04071518569628607),
    (11, 0.876942461150777, 0.04085918281634371),
    (11.1, 0.8817843643127137, 0.041321173576528536),
    (11.2, 0.8866262674746506, 0.04178316433671325),
    (11.3, 0.8914681706365872, 0.04224515509689808),
    (11.4, 0.8963100737985241, 0.042707145857082796),
    (11.5, 0.9011519769604608, 0.04316913661726762),
    (11.6, 0.9059938801223975, 0.04363112737745245),
    (11.7, 0.9108357832843343, 0.044093118137637166),
    (11.8, 0.915677686446271, 0.04455510889782199),
    (11.9, 0.9205195896082079, 0.04501709965800682),
    (12, 0.9253614927701446, 0.04547909041819176),
    (12.1, 0.9284214315713685, 0.04513109737805243),
    (12.2, 0.9314813703725926, 0.044783104337913215),
    (12.3, 0.9345413091738165, 0.04443511129777411),
    (12.4, 0.9376012479750406, 0.044087118257634894),
    (12.5, 0.9406611867762645, 0.04373912521749568),
    (12.6, 0.9437211255774884, 0.04339113217735646),
    (12.7, 0.9467810643787126, 0.043043139137217246),
    (12.8, 0.9498410031799365, 0.04269514609707803),
    (12.9, 0.9529009419811606, 0.04234715305693881),
    (13, 0.9559608807823845, 0.04199916001679971),
    (13.1, 0.95862482750345, 0.04181916361672766),
    (13.2, 0.9612887742245156, 0.041639167216655726),
    (13.3, 0.9639527209455812, 0.04145917081658368),
    (13.4, 0.9666166676666469, 0.04127917441651163),
    (13.5, 0.9692806143877124, 0.041099178016439586),
    (13.6, 0.9719445611087779, 0.04091918161636754),
    (13.7, 0.9746085078298435, 0.040739185216295604),
    (13.8, 0.977272454550909, 0.04055918881622356),
    (13.9, 0.9799364012719748, 0.04037919241615151),
    (14, 0.9826003479930403, 0.040199196016079686),
    (14.1, 0.9833203335933283, 0.039785204295914145),
    (14.2, 0.9840403191936163, 0.03937121257574849),
    (14.3, 0.9847603047939042, 0.03895722085558284),
    (14.4, 0.9854802903941922, 0.0385432291354173),
    (14.5, 0.9862002759944802, 0.03812923741525165),
    (14.6, 0.9869202615947682, 0.03771524569508611),
    (14.7, 0.9876402471950563, 0.037301253974920456),
    (14.8, 0.9883602327953441, 0.036887262254754805),
    (14.9, 0.9890802183956321, 0.036473270534589486),
    (15, 0.9898002039959202, 0.03605927881442361),
    (15.1, 0.9908201835963281, 0.036101277974440404),
    (15.2, 0.9918401631967362, 0.0361432771344572),
    (15.3, 0.9928601427971441, 0.03618527629447421),
    (15.4, 0.9938801223975522, 0.036227275454491004),
    (15.5, 0.9949001019979601, 0.036269274614507574),
    (15.6, 0.995920081598368, 0.03631127377452459),
    (15.7, 0.9969400611987761, 0.03635327293454138),
    (15.8, 0.997960040799184, 0.03639527209455817),
    (15.9, 0.9989800203995921, 0.036437271254574966),
    (16, 1, 0.03647927041459176),
]



def _build_kr_coeffs(raw_rows):
    """Muuntaa Excelin imperial-kertoimet cm/kg-yksiköihin (kuten Excel)."""
    return {
        age: (b0_in * 2.54, height_coef, weight_lb * 5.599741, mid_parent_coef)
        for age, b0_in, height_coef, weight_lb, mid_parent_coef in raw_rows
    }


def _build_reference(raw_rows):
    ages = [r[0] for r in raw_rows]
    means = [r[1] for r in raw_rows]
    sds = [r[2] for r in raw_rows]
    by_age = {
        round(a, 1): (m, s)
        for a, m, s in raw_rows
        if abs(a * 10 - round(a * 10)) < 1e-9
    }
    return {"ages": ages, "means": means, "sds": sds, "by_age": by_age}


KR_COEFFS = {
    "Poika": _build_kr_coeffs(_KR_COEFFS_BOYS_RAW),
    "Tyttö": _build_kr_coeffs(_KR_COEFFS_GIRLS_RAW),
}
SWE_REF = {
    "Poika": _build_reference(_SWE_REF_BOYS_RAW),
    "Tyttö": _build_reference(_SWE_REF_GIRLS_RAW),
}


# ---------------------------------------------------------------------------
# Laskenta
# ---------------------------------------------------------------------------
def _is_last_day_of_february(d):
    return d.month == 2 and (d + datetime.timedelta(days=1)).month == 3


def yearfrac(d1, d2):
    """Excelin YEARFRAC(d1, d2) oletusperusteella 0 (US 30/360)."""
    if d1 > d2:
        d1, d2 = d2, d1
    day1, day2 = d1.day, d2.day
    if _is_last_day_of_february(d1) and _is_last_day_of_february(d2):
        day2 = 30
    if _is_last_day_of_february(d1):
        day1 = 30
    if day2 == 31 and day1 >= 30:
        day2 = 30
    if day1 == 31:
        day1 = 30
    days360 = (d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + (day2 - day1)
    return days360 / 360


def excel_round(value, digits=0):
    """Excelin ROUND: puolikkaat pyöristetään ylöspäin (Pythonin round ei tee niin)."""
    quantum = Decimal(1).scaleb(-digits)
    return float(Decimal(repr(round(value, 10))).quantize(quantum, rounding=ROUND_HALF_UP))


def mround(value, multiple):
    """Excelin MROUND positiivisille luvuille (puolikkaat pyöristetään ylös)."""
    return math.floor(value / multiple + 0.5) * multiple


def average(a, b):
    return (a + b) / 2


def calc_mirwald(gender, dob, dom, height_cm, weight_kg, sitting_height_cm):
    """Mirwald: maturity offset ja PHV-ikä (Excel: "Mirwald PHV ikä").

    height_cm, weight_kg ja sitting_height_cm ovat kahden mittauksen keskiarvoja.
    sitting_height_cm on istumapituus ilman penkkiä (penkin korkeus vähennetty).
    """
    age = yearfrac(dob, dom)
    leg_length = height_cm - sitting_height_cm
    leg_x_sit = leg_length * sitting_height_cm
    age_x_leg = age * leg_length
    age_x_sit = age * sitting_height_cm
    age_x_weight = age * weight_kg
    weight_height_ratio = weight_kg / height_cm * 100

    if gender == "Tyttö":
        offset = (
            MIRWALD_GIRLS_INTERCEPT
            + 0.0001882 * leg_x_sit
            + 0.0022 * age_x_leg
            + 0.005841 * age_x_sit
            - 0.002658 * age_x_weight
            + 0.07693 * weight_height_ratio
        )
    else:
        offset = (
            MIRWALD_BOYS_INTERCEPT
            + 0.0002708 * leg_x_sit
            - 0.001663 * age_x_leg
            + 0.007216 * age_x_sit
            + 0.02292 * weight_height_ratio
        )

    phv_age = age - offset
    limit = MIRWALD_LIMITS[gender][dob.month - 1]
    margin = phv_age - limit
    eligible = phv_age > limit
    return {
        "age": age,
        "height": height_cm,
        "weight": weight_kg,
        "sitting_height": sitting_height_cm,
        "leg_length": leg_length,
        "offset": offset,
        "phv_age": phv_age,
        "limit": limit,
        "margin": margin,
        "eligible": eligible,
        "result": "KYLLÄ" if eligible else "EI",
    }


def maturity_status(pct):
    """Kypsyysvaihe %PAH:n mukaan (Excel: Maturity status)."""
    if pct < 0.86:
        return "Pre PHV"
    if pct < 0.90:
        return "Nearing PHV"
    if pct < 0.92:
        return "At PHV"
    if pct < 0.95:
        return "Departing PHV"
    return "Post PHV"


def bio_timing(age_offset):
    """Biologinen ajoitus (Excel: Bio timing)."""
    if age_offset > 0.5:
        return "EARLY"
    if age_offset < -0.5:
        return "LATE"
    return "ON-TIME"


def calc_khamis_roche(gender, dob, dom, height_cm, weight_kg, mum_cm, dad_cm, age_group):
    """Khamis-Roche: PAH, %PAH, biologinen ikä ja sääntö (Excel: BOYS / GIRLS)."""
    age = yearfrac(dob, dom)
    rounded_age = mround((dom - dob).days / 365.25, 0.5)
    coeffs = KR_COEFFS[gender].get(rounded_age)
    if coeffs is None:
        raise ValueError(
            f"Pelaajan ikä ({age:.1f} v) on Khamis-Roche-kertoimien alueen "
            "(4–17,5 v) ulkopuolella."
        )
    intercept, height_coef, weight_coef, mid_parent_coef = coeffs

    # Vanhempien pituudet korjataan (Excel: Hidden_boys / Hidden_girls, sarakkeet K-Q).
    mum_adj_cm = (2.803 + 0.953 * (mum_cm / 2.54)) * 2.54
    dad_adj_cm = (2.316 + 0.955 * (dad_cm / 2.54)) * 2.54
    mid_parent_cm = (mum_adj_cm + dad_adj_cm) / 2

    pah = intercept + height_coef * height_cm + weight_coef * weight_kg + mid_parent_coef * mid_parent_cm
    if pah < 100:
        raise ValueError(f"Laskettu PAH ({pah:.1f} cm) on epärealistinen. Tarkista syötetyt mitat.")

    pct_pah = height_cm / pah
    if gender == "Tyttö":
        pct_pah = min(pct_pah, 1.0)  # Excel rajaa tytöillä %PAH:n 100 %:iin

    # Biologinen ikä: viiteaineiston ikä, jonka %PAH on lähimpänä pelaajan arvoa.
    ref = SWE_REF[gender]
    best = min(range(len(ref["means"])), key=lambda i: abs(ref["means"][i] - pct_pah))
    bio_age = ref["ages"][best]

    age_offset = bio_age - age

    z_score = None
    ref_point = ref["by_age"].get(excel_round(age, 1))
    if ref_point is not None:
        mean, sd = ref_point
        z_score = (height_cm / pah * 100 - mean * 100) / (sd * 100)

    limit = KR_LIMITS[age_group][dob.month - 1]
    margin = limit - bio_age
    eligible = bio_age < limit
    return {
        "age": age,
        "rounded_age": rounded_age,
        "height": height_cm,
        "weight": weight_kg,
        "mid_parent": mid_parent_cm,
        "pah": pah,
        "pct_pah": pct_pah,
        "status": maturity_status(pct_pah),
        "bio_age": bio_age,
        "age_offset": age_offset,
        "timing": bio_timing(age_offset),
        "z_score": z_score,
        "limit": limit,
        "margin": margin,
        "eligible": eligible,
        "result": "KYLLÄ" if eligible else "EI",
    }


# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------
def load_credentials_info():
    """Lukee palvelutilin tiedot Secretsistä, tukee kahta tallennustapaa."""
    if "gcp_service_account" in st.secrets:
        return dict(st.secrets["gcp_service_account"])  # TOML-taulukko
    if "GCP_SERVICE_ACCOUNT" in st.secrets:
        return json.loads(st.secrets["GCP_SERVICE_ACCOUNT"])  # JSON merkkijonona
    raise KeyError("Secretsistä puuttuu 'gcp_service_account' tai 'GCP_SERVICE_ACCOUNT'.")


def add_months(d, months):
    """Lisää päivämäärään kalenterikuukausia (kuukauden lopussa päivä rajataan)."""
    idx = d.year * 12 + (d.month - 1) + months
    year, month = divmod(idx, 12)
    month += 1
    last_day = (datetime.date(year + (month == 12), month % 12 + 1, 1) - datetime.timedelta(days=1)).day
    return datetime.date(year, month, min(d.day, last_day))


def sheet_safe(value):
    """Estää kaavainjektion: tekstin alussa oleva = + - @ muutetaan tekstiksi."""
    if isinstance(value, str) and value[:1] in ("=", "+", "-", "@"):
        return "'" + value
    return value


def save_to_google_sheets(row_data):
    """Tallentaa rivin taulukkoon. Palauttaa (onnistui, virhe, varoitus)."""
    try:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        credentials = Credentials.from_service_account_info(load_credentials_info(), scopes=scopes)
        client = gspread.authorize(credentials)

        # Jos Secretsissä on SHEET_ID, taulukko avataan sillä (ohittaa nimihaun).
        if "SHEET_ID" in st.secrets:
            spreadsheet = client.open_by_key(st.secrets["SHEET_ID"])
        else:
            spreadsheet = client.open(SHEET_NAME)
        sheet = spreadsheet.sheet1

        warning = None
        existing_header = [h.strip() for h in sheet.row_values(1)]
        if not existing_header:
            sheet.insert_row(SHEET_HEADERS, 1)  # tyhjä taulukko: lisätään otsikot
        elif existing_header[: len(SHEET_HEADERS)] != SHEET_HEADERS:
            warning = (
                "Taulukon otsikkorivi ei vastaa sovelluksen tallentamia sarakkeita. "
                f"Korvaa rivin 1 otsikot näillä (sarakkeet A–{chr(64 + len(SHEET_HEADERS))}): " + " | ".join(SHEET_HEADERS)
            )

        sheet.append_row(row_data, value_input_option="USER_ENTERED")
        return True, None, warning
    except Exception as e:  # noqa: BLE001
        return False, f"{type(e).__name__}: {e}\n\n{traceback.format_exc()}", None


# ---------------------------------------------------------------------------
# Käyttöliittymä
# ---------------------------------------------------------------------------
def render_mirwald_result(res, gender, pallo_id, month_name):
    st.success(f"Mirwald-tulokset laskettu PalloID:lle {pallo_id}:")
    r1, r2, r3 = st.columns(3)
    r1.metric("Kronologinen ikä", f"{res['age']:.2f} v")
    r2.metric("Arvioitu PHV-ikä", f"{res['phv_age']:.2f} v")
    r3.metric(
        "Pelioikeudellinen (Yli-ikäisyyssääntö)",
        res["result"],
        delta="Täyttää vaatimuksen" if res["eligible"] else "Ei täytä vaatimusta",
        delta_color="normal" if res["eligible"] else "inverse",
    )

    st.caption(
        f"Käytetty penkin korkeus: {res['bench']:.1f} cm (vähennetty istumapituusmittauksista). "
        "Tarkista, että arvo vastaa käyttämääsi penkkiä."
    )

    group = "pojat" if gender == "Poika" else "tytöt"
    if res["eligible"]:
        st.success(
            f"**Tulkinta: KYLLÄ.** PHV-ikä {res['phv_age']:.2f} v **ylittää** raja-arvon "
            f"{res['limit']:.2f} v ({group}, syntymäkuukausi {month_name}) {res['margin']:.2f} vuodella."
        )
    else:
        st.warning(
            f"**Tulkinta: EI.** PHV-ikä {res['phv_age']:.2f} v **ei ylitä** raja-arvoa "
            f"{res['limit']:.2f} v ({group}, syntymäkuukausi {month_name}). Puuttuu {-res['margin']:.2f} vuotta."
        )


def render_kr_result(res, pallo_id, age_group, month_name):
    st.success(f"Khamis-Roche-tulokset laskettu PalloID:lle {pallo_id}:")
    r1, r2, r3 = st.columns(3)
    r1.metric("Kronologinen ikä", f"{res['age']:.2f} v")
    r2.metric("Biologinen ikä", f"{res['bio_age']:.1f} v")
    r3.metric(
        "Pelioikeudellinen (Yli-ikäisyyssääntö)",
        res["result"],
        delta="Täyttää vaatimuksen" if res["eligible"] else "Ei täytä vaatimusta",
        delta_color="normal" if res["eligible"] else "inverse",
    )
    r4, r5, r6 = st.columns(3)
    r4.metric("Ennustettu aikuispituus (PAH)", f"{res['pah']:.1f} cm")
    r5.metric("%PAH", f"{res['pct_pah'] * 100:.1f} %")
    r6.metric("Maturity status", res["status"])

    if res["eligible"]:
        st.success(
            f"**Tulkinta: KYLLÄ.** Biologinen ikä {res['bio_age']:.1f} v **alittaa** raja-arvon "
            f"{res['limit']:.2f} v ({age_group}v, syntymäkuukausi {month_name}) {res['margin']:.2f} vuodella."
        )
    else:
        st.warning(
            f"**Tulkinta: EI.** Biologinen ikä {res['bio_age']:.1f} v **ei alita** raja-arvoa "
            f"{res['limit']:.2f} v ({age_group}v, syntymäkuukausi {month_name}). Ylitys {-res['margin']:.2f} vuotta."
        )


def render_mirwald_instructions():
    st.markdown(
        """
**Mirwald-menetelmä**

- **Maturity offset** kertoo, kuinka monta vuotta pelaajan kasvupyrähdyksen huippukohdasta
  (PHV, Peak Height Velocity) on kulunut. Miinusmerkkinen arvo tarkoittaa, että huippukohta on
  vielä edessä, plusmerkkinen että se on jo ohitettu.
- **PHV-ikä** = kronologinen ikä − maturity offset, eli arvioitu ikä, jolloin kasvupyrähdys on
  voimakkaimmillaan. Mitä suurempi PHV-ikä, sitä myöhemmin kasvupyrähdyksen arvioidaan tulevan.
- **Sääntö:** pelaajan PHV-iän tulee **ylittää** syntymäkuukautta vastaava raja-arvo.
  Pojilla ja tytöillä on omat raja-arvonsa.
- **Istumapituus:** mittaa istumapituus penkiltä. Sovellus vähentää penkin korkeuden
  (oletus 40 cm) mittaustuloksesta. Jos mittaat istumapituuden suoraan ilman penkkiä,
  aseta penkin korkeudeksi 0.
"""
    )
    with st.expander("Mirwald: raja-arvot syntymäkuukauden mukaan (PHV-iän tulee ylittää arvo)"):
        table = pd.DataFrame(
            {
                "Syntymäkuukausi": MONTHS_FI,
                "Pojat": [f"{v:.2f}" for v in MIRWALD_LIMITS["Poika"]],
                "Tytöt": [f"{v:.2f}" for v in MIRWALD_LIMITS["Tyttö"]],
            }
        )
        st.dataframe(table, hide_index=True)
    with st.expander("Mirwald: esimerkit"):
        st.markdown(
            """
- Poika, syntymäkuukausi syyskuu, PHV-ikä 14,6 → raja-arvo 14,30 → **KYLLÄ**
- Poika, syntymäkuukausi tammikuu, PHV-ikä 14,5 → raja-arvo 14,97 → **EI**
- Tyttö, syntymäkuukausi huhtikuu, PHV-ikä 12,2 → raja-arvo 12,82 → **EI**
- Tyttö, syntymäkuukausi lokakuu, PHV-ikä 12,5 → raja-arvo 12,32 → **KYLLÄ**
"""
        )


def render_kr_instructions():
    st.markdown(
        """
**Khamis-Roche-menetelmä**

- **PAH** (Predicted Adult Height) on ennustettu aikuispituus. Se lasketaan pelaajan pituudesta,
  painosta ja vanhempien pituuksista (ikäkohtaisilla kertoimilla).
- **%PAH** = pelaajan nykyinen pituus prosentteina ennustetusta aikuispituudesta.
- **Maturity status** (%PAH:n mukaan): Pre PHV (alle 86 %) · Nearing PHV (86–90 %) ·
  At PHV (90–92 %) · Departing PHV (92–95 %) · Post PHV (95 % tai yli).
- **Biologinen ikä** on ruotsalaisen viiteaineiston (Albertsson ym.) ikä, jonka keskimääräinen
  %PAH on lähimpänä pelaajan %PAH-arvoa.
- **Ikäero (Age offset)** = biologinen ikä − kronologinen ikä. **Bio timing:** yli +0,5 v = EARLY
  (aikainen), −0,5…+0,5 v = ON-TIME, alle −0,5 v = LATE (myöhäinen).
- **Z-score** kertoo, montako keskihajontaa pelaajan %PAH poikkeaa viiteaineiston keskiarvosta
  samassa iässä.
- **Sääntö:** pelaajan biologisen iän tulee **alittaa** ikäryhmän ja syntymäkuukauden mukainen
  raja-arvo.
- **Vanhempien pituudet:** syötä biologisten vanhempien pituudet senttimetreinä.
"""
    )
    with st.expander(
        "Khamis-Roche: raja-arvot ikäryhmän ja syntymäkuukauden mukaan (biologisen iän tulee alittaa arvo)"
    ):
        st.caption("Arvo = ikäryhmän keskimääräinen kronologinen ikä vähennettynä 0,75 vuodella.")
        table = pd.DataFrame(
            {
                "Syntymäkuukausi": MONTHS_FI,
                "13v": [f"{v:.2f}" for v in KR_LIMITS[13]],
                "14v": [f"{v:.2f}" for v in KR_LIMITS[14]],
                "15v": [f"{v:.2f}" for v in KR_LIMITS[15]],
            }
        )
        st.dataframe(table, hide_index=True)
    with st.expander("Khamis-Roche: esimerkit"):
        st.markdown(
            """
- Syntymäkuukausi syyskuu, ikäryhmä 13v, biologinen ikä 13,0 → raja-arvo 12,48 → **EI**
- Syntymäkuukausi tammikuu, ikäryhmä 14v, biologinen ikä 12,6 → raja-arvo 12,75 → **KYLLÄ**
"""
        )


def render_instructions(is_mirwald):
    st.divider()
    st.header("📖 Ohjeet tuloksen tulkintaan")
    st.markdown(
        """
**Näin tulos syntyy**

1. Mittaa pelaajan pituus ja paino kaksi kertaa (Mirwald-menetelmässä myös istumapituus).
   Sovellus laskee mittausten keskiarvot.
2. Sovellus laskee pelaajan biologisen kypsyyden valitulla menetelmällä.
3. Tulosta verrataan pelaajan **syntymäkuukautta** vastaavaan raja-arvoon.
4. **KYLLÄ** = pelaaja täyttää yli-ikäisyyssäännön vaatimuksen.
   **EI** = pelaaja ei täytä vaatimusta.
"""
    )

    # Molempien menetelmien ohjeet ovat aina näkyvissä; valittu menetelmä on ensimmäisenä.
    if is_mirwald:
        tab_mirwald, tab_kr = st.tabs(["Mirwald", "Khamis-Roche"])
    else:
        tab_kr, tab_mirwald = st.tabs(["Khamis-Roche", "Mirwald"])
    with tab_mirwald:
        render_mirwald_instructions()
    with tab_kr:
        render_kr_instructions()

    st.markdown(
        """
**Huomioitavaa**

- Tulos on laskennallinen arvio mittausten perusteella. Jos ero raja-arvoon on pieni,
  pienikin mittausvirhe voi muuttaa tuloksen, joten tarkista mittaukset huolellisesti.
- Jos kaksi mittausta poikkeavat toisistaan selvästi, mittaa uudelleen.
"""
    )


HERO_IMAGE_NAMES = ("hero.jpg", "hero.jpeg", "hero.png", "hero.webp")
HERO_MIME = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png", "webp": "image/webp"}


def hero_image_css():
    """Palauttaa taustakuvan CSS:n, jos repossa on hero.jpg/.png/.webp app.py:n vieressä."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for name in HERO_IMAGE_NAMES:
        path = os.path.join(base_dir, name)
        if os.path.isfile(path):
            try:
                with open(path, "rb") as f:
                    encoded = base64.b64encode(f.read()).decode("ascii")
            except OSError:
                return None
            mime = HERO_MIME[name.rsplit(".", 1)[1]]
            return f"url(data:{mime};base64,{encoded})"
    return None


def render_hero():
    """Sivun yläosan banneri: sininen sävytetty taustakuva (tai pelkkä sininen tausta) ja otsikko."""
    image = hero_image_css()
    overlay = "linear-gradient(rgba(8, 48, 140, 0.84), rgba(8, 48, 140, 0.84))"
    background = f"{overlay}, {image} center / cover" if image else "linear-gradient(135deg, #0b3a96, #06246b)"
    st.markdown(
        f"""
        <style>
        .pb-hero {{
            background: {background};
            border-radius: 14px;
            padding: 44px 20px 36px 20px;
            margin-bottom: 1.2rem;
            text-align: center;
            color: #ffffff;
        }}
        .pb-hero-title {{
            font-family: "Trebuchet MS", "Segoe UI", Arial, sans-serif;
            font-style: italic;
            font-weight: 800;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            font-size: clamp(1.6rem, 6vw, 2.6rem);
            line-height: 1.1;
            margin: 0;
        }}
        .pb-hero-title span {{ color: #f2a99b; }}
        .pb-hero-sub {{
            font-size: clamp(0.95rem, 2.6vw, 1.15rem);
            margin: 14px auto 0 auto;
            max-width: 560px;
            opacity: 0.95;
        }}
        </style>
        <div class="pb-hero">
            <div class="pb-hero-title">Biologinen ikä <span>&amp; pelioikeus</span></div>
            <div class="pb-hero-sub">Palloliiton yli-ikäisyyssäännön tarkistus ja ilmoitus Palloliitolle</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_history():
    """Seuran omat tulokset: istunnon mittaukset ladattavana tiedostona."""
    history = st.session_state.get("history", [])
    if not history:
        return
    st.subheader("Seuran omat tulokset (tämä istunto)")
    st.caption(
        "Lataa tulokset omaan käyttöön ja säilytä ne seuran järjestelmässä. Lista tyhjenee, kun sivu ladataan "
        "uudelleen tai istunto päättyy."
    )
    df = pd.DataFrame(history)
    st.dataframe(df, hide_index=True)
    csv = df.to_csv(index=False, sep=";").encode("utf-8-sig")  # ; ja BOM: aukeaa suoraan suomalaisessa Excelissä
    st.download_button(
        "Lataa tulokset (CSV)",
        data=csv,
        file_name=f"palloliitto_bio_{datetime.date.today().isoformat()}.csv",
        mime="text/csv",
    )


def main():
    st.set_page_config(page_title="Palloliitto Bio", page_icon="⚽", layout="centered")

    render_hero()
    st.write(
        "Valitse testausmenetelmä, syötä pelaajan 8-numeroinen PalloID, syntymäaika ja kaksi "
        "mittausta kustakin mitasta. Sovellus laskee keskiarvot ja tarkistaa yli-ikäisyyssäännön "
        "automaattisesti. Tulos ilmoitetaan samalla Palloliitolle, ja seura voi ladata omat tuloksensa "
        "sivun alalaidasta."
    )

    test_type = st.radio(
        "Valitse testausmenetelmä:",
        [
            "Mirwald (PHV & Maturity Offset)",
            "Khamis-Roche (Aikuisiän pituusennuste / PAH)",
        ],
    )
    is_mirwald = "Mirwald" in test_type

    with st.form("bio_form"):
        st.subheader("1. Seuran ja testaajan tiedot")
        club = st.text_input("Seura", placeholder="Esim. FC Kontu")
        ccol1, ccol2 = st.columns(2)
        with ccol1:
            tester_name = st.text_input("Testaajan nimi")
        with ccol2:
            tester_email = st.text_input("Testaajan sähköposti")

        st.subheader("2. Pelaajan tunnistetiedot")
        col1, col2 = st.columns(2)
        with col1:
            pallo_id = st.text_input("PalloID (8 numeroa)", max_chars=8, placeholder="Esim. 12345678")
            gender = st.selectbox("Sukupuoli", ["Poika", "Tyttö"])
        with col2:
            dob = st.date_input(
                "Syntymäaika",
                value=datetime.date(2012, 1, 1),
                min_value=datetime.date(2005, 1, 1),
                max_value=datetime.date(2022, 12, 31),
                format="DD.MM.YYYY",
            )
            dom = st.date_input("Mittauspäivä", value=datetime.date.today(), format="DD.MM.YYYY")

        st.subheader("3. Mittaukset")
        st.caption("Mittaa jokainen mitta kaksi kertaa. Sovellus käyttää mittausten keskiarvoa.")
        mcol1, mcol2 = st.columns(2)
        with mcol1:
            height_1 = st.number_input(
                "Seisomapituus, mittaus 1 (cm)", min_value=100.0, max_value=220.0, value=150.0, step=0.1
            )
            weight_1 = st.number_input(
                "Paino, mittaus 1 (kg)", min_value=20.0, max_value=120.0, value=40.0, step=0.1
            )
            if is_mirwald:
                sitting_1 = st.number_input(
                    "Istumapituus, mittaus 1 (cm)", min_value=50.0, max_value=160.0, value=115.0, step=0.1
                )
        with mcol2:
            height_2 = st.number_input(
                "Seisomapituus, mittaus 2 (cm)", min_value=100.0, max_value=220.0, value=150.0, step=0.1
            )
            weight_2 = st.number_input(
                "Paino, mittaus 2 (kg)", min_value=20.0, max_value=120.0, value=40.0, step=0.1
            )
            if is_mirwald:
                sitting_2 = st.number_input(
                    "Istumapituus, mittaus 2 (cm)", min_value=50.0, max_value=160.0, value=115.0, step=0.1
                )

        if is_mirwald:
            bench_height = st.number_input(
                "Penkin korkeus (cm)",
                min_value=0.0,
                max_value=80.0,
                value=DEFAULT_BENCH_HEIGHT_CM,
                step=0.5,
                help=(
                    "Penkin korkeus vähennetään istumapituusmittauksesta (Excel: 40 cm). "
                    "Jos mittaat istumapituuden suoraan ilman penkkiä, aseta 0."
                ),
            )
            mum_height = dad_height = None
            kr_group = None
        else:
            pcol1, pcol2 = st.columns(2)
            with pcol1:
                mum_height = st.number_input(
                    "Äidin pituus (cm)", min_value=130.0, max_value=200.0, value=165.0, step=0.5
                )
            with pcol2:
                dad_height = st.number_input(
                    "Isän pituus (cm)", min_value=140.0, max_value=220.0, value=180.0, step=0.5
                )
            kr_group = st.selectbox("Ikäryhmä (Khamis-Roche)", [13, 14, 15])

        submitted = st.form_submit_button("Laske ja tarkista pelioikeus")

    if submitted:
        pallo_id = pallo_id.strip()
        club = club.strip()
        tester_name = tester_name.strip()
        tester_email = tester_email.strip()
        if not club or not tester_name:
            st.warning("Syötä seura ja testaajan nimi.")
        elif not EMAIL_RE.match(tester_email):
            st.warning("Syötä testaajan sähköposti oikeassa muodossa (esim. nimi@seura.fi).")
        elif not pallo_id or len(pallo_id) != 8 or not pallo_id.isdigit():
            st.warning("Syötä kelvollinen 8-numeroinen PalloID.")
        elif dom <= dob:
            st.warning("Mittauspäivän tulee olla syntymäajan jälkeen.")
        else:
            height = average(height_1, height_2)
            weight = average(weight_1, weight_2)
            try:
                if is_mirwald:
                    sitting = average(sitting_1, sitting_2) - bench_height
                    if sitting <= 0 or sitting >= height:
                        raise ValueError(
                            "Istumapituus (penkin korkeus vähennettynä) on epärealistinen. "
                            "Tarkista istumapituus ja penkin korkeus."
                        )
                    res = calc_mirwald(gender, dob, dom, height, weight, sitting)
                    res["bench"] = bench_height
                else:
                    res = calc_khamis_roche(gender, dob, dom, height, weight, mum_height, dad_height, kr_group)
            except ValueError as err:
                st.error(str(err))
                res = None

            if res is not None:
                month_name = MONTHS_FI[dob.month - 1].lower()
                if is_mirwald:
                    render_mirwald_result(res, gender, pallo_id, month_name)
                else:
                    render_kr_result(res, pallo_id, kr_group, month_name)

                with st.expander("Laskennan yksityiskohdat"):
                    details = [("Mittauspäivän ja syntymäajan välinen ikä (v)", f"{res['age']:.3f}")]
                    details += [
                        ("Pituus, keskiarvo (cm)", f"{res['height']:.2f}"),
                        ("Paino, keskiarvo (kg)", f"{res['weight']:.2f}"),
                    ]
                    if is_mirwald:
                        details += [
                            ("Penkin korkeus (cm)", f"{res['bench']:.1f}"),
                            ("Istumapituus ilman penkkiä, keskiarvo (cm)", f"{res['sitting_height']:.2f}"),
                            ("Jalkojen pituus (cm)", f"{res['leg_length']:.2f}"),
                            ("Maturity offset (v)", f"{res['offset']:.3f}"),
                            ("PHV-ikä (v)", f"{res['phv_age']:.3f}"),
                        ]
                    else:
                        details += [
                            ("Korjattu vanhempien keskipituus (cm)", f"{res['mid_parent']:.2f}"),
                            ("Pyöristetty ikä kertoimia varten (v)", f"{res['rounded_age']:.1f}"),
                            ("PAH (cm)", f"{res['pah']:.2f}"),
                            ("%PAH", f"{res['pct_pah'] * 100:.2f} %"),
                            ("Biologinen ikä (v)", f"{res['bio_age']:.1f}"),
                            ("Ikäero, biologinen − kronologinen (v)", f"{res['age_offset']:.2f}"),
                            ("Bio timing", res["timing"]),
                            ("Z-score", "–" if res["z_score"] is None else f"{res['z_score']:.2f}"),
                        ]
                    details.append(("Raja-arvo (v)", f"{res['limit']:.2f}"))
                    st.table(pd.DataFrame(details, columns=["Suure", "Arvo"]).set_index("Suure"))

                valid_until = add_months(dom, VALIDITY_MONTHS) if res["eligible"] else None
                record = {
                    "Pvm": dom.strftime("%d.%m.%Y"),
                    "Seura": club,
                    "Testaaja": tester_name,
                    "Testaajan sähköposti": tester_email,
                    "Metodi": "Mirwald" if is_mirwald else "Khamis-Roche",
                    "PalloID": pallo_id,
                    "Sukupuoli": gender,
                    "Syntymäkuukausi": dob.month,
                    "Ikäryhmä (K-R)": "" if is_mirwald else kr_group,
                    "Ikä": round(res["age"], 2),
                    "Pituus ka (cm)": round(res["height"], 2),
                    "Paino ka (kg)": round(res["weight"], 2),
                    "Istumapituus ka (cm)": round(res["sitting_height"], 2) if is_mirwald else "",
                    "Penkin korkeus (cm)": round(res["bench"], 1) if is_mirwald else "",
                    "PHV-ikä / Bio-ikä": round(res["phv_age"] if is_mirwald else res["bio_age"], 2),
                    "Raja-arvo": round(res["limit"], 2),
                    "Pelioikeus": res["result"],
                    "Mittaus voimassa asti": valid_until.strftime("%d.%m.%Y") if valid_until else "",
                }
                row_to_save = []
                for header in SHEET_HEADERS:
                    value = record[header]
                    if header == "PalloID":
                        value = "'" + value  # heittomerkki säilyttää etunollat tekstinä
                    row_to_save.append(sheet_safe(value))

                ok, error_message, warning = save_to_google_sheets(row_to_save)
                if ok:
                    st.success("Tulos on ilmoitettu Palloliitolle ja tallennettu.")
                    if warning:
                        st.warning(warning)
                else:
                    st.error(
                        "Ilmoitus Palloliitolle (tallennus Google Sheetsiin) epäonnistui. "
                        "Lataa tulos omaan käyttöön alta ja yritä myöhemmin uudelleen."
                    )
                    st.code(error_message)

                if res["eligible"]:
                    st.info(
                        f"**Mittaustulos on voimassa {VALIDITY_MONTHS} kuukautta mittauspäivästä, "
                        f"eli {valid_until.strftime('%d.%m.%Y')} asti.** Seura vastaa siitä, että voimassa oleva "
                        "mittaustulos on käytettävissä koko pelioikeuden voimassaolon ajan ja että uusi mittaus "
                        "tehdään tarvittaessa ennen voimassaoloajan päättymistä."
                    )
                else:
                    st.caption("Pelioikeutta ei synny tämän mittauksen perusteella, joten voimassaoloaikaa ei ole.")

                history = st.session_state.setdefault("history", [])
                record_view = dict(record)
                record_view["Ilmoitus Palloliitolle"] = "Lähetetty" if ok else "EI lähetetty"
                history.append(record_view)

    render_history()

    render_instructions(is_mirwald)


if __name__ == "__main__":
    main()
