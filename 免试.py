# !/usr/bin/env python
# -*- coding: utf-8 -*-
# ------------------------------
''''''
import time
from charset_normalizer import cd
from openpyxl import Workbook
import requests
from bs4 import BeautifulSoup
from lxml import etree
import re
import json
import os
import requests
wb = Workbook()
sheet = wb.active
sheet.append(['姓名','自我介绍？','国家','渴望薪资','工作成功率','赚了多少','自我介绍','第一技能'])
a = []
b = []
c = []
d = []
e = []
f = []
g = []
h = []
i = []
pages = int(input('想要多少页自己填：'))
      #2500个对我的小电脑来说有点吃力了，希望看到这段代码的学姐学长谅解，就做成这种了。
cookies = {
    'cdeb28a6sb': 'oauth2v2_int_a9eeb64158476b9c19485d065378f0c8',
    'visitor_id': '223.112.31.154.1790168435779000',
    '__cflb': '02DiuEXPXZVk436fJfSVuuwDqLqkhavJbxwg9CSZdKfEB',
    'fp_token_7c6a6574-f011-4c9a-abdd-9894a102ccef': 'e7z7UM0kkn51kR5sbNcqpbxUKszIKupoRHgU7vej3hU=',
    '_cq_duid': '1.1790168441.b7inZu6F2tLrsTYQ',
    'spt': '4b66fc4d-6fc9-412d-b0f6-3297a0cc0ac1',
    '__ps_r': '_',
    '__ps_lu': 'https://www.upwork.com/',
    '__ps_did': 'pscrb_832e8a34-0258-4b9f-9512-e248ba0dbea0',
    '__ps_fva': '1790168463522',
    '__pdst': '44722a29abfa42258dd7e8457bb27601',
    '_fbp': 'fb.1.1790168465552.734314235541176021',
    'tatari-session-cookie': 'e135e25e-0891-2436-31e9-28bf86e89a38',
    '_ga': 'GA1.1.1926171419.1790168464',
    '_tt_enable_cookie': '1',
    '_ttp': '01M377VFRM11E9JE4JMJJ433BV_.tt.1.1790170742549',
    'recognized': '6e3cb5ff0eddd1b2',
    'current_organization_uid': '2102754721647967795',
    'company_last_accessed': '1020210583',
    'ngIaSidebar': 'o.0*p.*w.0*e.1',
    'uThemeNull': '1',
    'tatari-user-cookie': '2102754721647967794',
    'device_view': 'full',
    'g_state': '{"i_l":0,"i_ll":1790175734586,"i_b":"fZGgeg3w+R/FWn9wWA6tsXqIwPIFM2GIYzN71ezzF4E","i_e":{"enable_itp_optimization":24},"i_et":1790175734586}',
    'console_user': '6e3cb5ff0eddd1b2',
    'enabled_ff': '!CI12577UniversalSearch,!Fluid,!MP16400Air3Migration,!OTBnrOn,!SFE1028UmaRecruiterRefineImprovements,!SSINavUser,!i18nGA,CI17409DarkModeUI,CmpLibOn,JPAir3,SFE919UmaRecruiterBasicQt3,SSINavUserBpa,TONB2256Air3Migration,TranscendUIOn,WP658TranscendOn,i18nOn',
    'cookie_prefix': '',
    'cookie_domain': '.upwork.com',
    'XSRF-TOKEN': '2102876886097338876',
    'master_access_token': '6fbe682b.oauth2v2_int_4da712c908797fe85ac5de87aaa7dbf1',
    'user_uid': '2102754721647967794',
    'oauth2_global_js_token': 'oauth2v2_int_f11a4e1288497c9fc6bfff7696a549a0',
    'fe54365bsb': 'oauth2v2_int_78532891cf6ebe551188885ea340e167',
    '_cq_suid': '1.1790199897.NFaGEAlDDO54ieYK',
    'IR_gbd': 'upwork.com',
    '_mkto_trk': 'id:518-RKL-392&token:_mch-upwork.com-e1c5aa472364fc0101c70fcfe2cab9af',
    'country_code': 'JP',
    'e53a11d0sb': 'oauth2v2_int_7a4052cd5516c763d3ab62afbf2d3020',
    'ftr_ncd': '6',
    'visitor_gql_token': 'oauth2v2_int_0c54b3723e80d7713c3afa84a961630f',
    '_cfuvid': 'VNc_KWhdyaT8X2l49WLCihXlR.QLNAFAneaKN4W.Tv0-1790264966.486338-1.0.1.1-4lWUxg86sudD0YgTDT5xm.8Ga..eVAFhXbQ72G3RLSk',
    'umq': '1699',
    'umq_cq': '1643',
    '_upw_ses.5831': '*',
    '__ps_sr': '_',
    '__ps_slu': 'https://www.upwork.com/nx/search/talent?nav_dir=pop&page=1&__cf_chl_tk=mB8kMyvbbaCRYWWyblNnCx6qRkvssoK4Y6N1WKAUE8Y-1790268907-1.0.1.1-AERFS49ht7XLSUccDPSAEQC0ewjbpyh_ZDdlWbA892Y',
    'profile_vv_gql_token': 'oauth2v2_int_031046fa0b29fec7bd96adaf2a4b16ba',
    'AttachmentsNuxt_vt': 'oauth2v2_int_c96df3d897b57f8b9c5b84f07af6d8a0',
    'UniversalSearchNuxt_vt': 'oauth2v2_int_dc47084c47186ee0161a359d2c6ced30',
    'cf_clearance': 'DUjRdG4sBPmUOglIbaWtK0a3YSNePla9uTWiUThhJlw-1790271700-1.2.1.1-SkgMHqvDiNd4x6TmMlj.JAx86kEZXTiBG2tI83vGN_JzSuXdLP81ZzPK6qegQas6SL44c.rcY3AltqWbTwFziAT6XGVbF6Crvu3PcsFVboe_a66PSn8BT69NedW9bHW2wXIrHfxk9p8SHZtOjbp9lddv7ZGIBEa7lPsVDMfm8L6L91N_gYwtGS7gnWqk9JOWaB5wz6GWfE9q5WKStuU1Kc0qiDLU5FgT8r9DWRYjPyZLbaeEbWriVzG89hwyeTCO4OsR.x4M.4TNZJYCVpsJ4.aymYex7WXU_XsHRlg3vaAoMdD3e9xgHeee4alKk8wObq3ktBFqOPVOuaSEjPOPA5V1ReWFCXLhHVunInyoEXKwkGifSS2JKY7BpqYSNl.i31E.1kecSx95FROxGkCut9MrQf6VwLVlUWWFbKsgi_WVy..pTvzppZNwbbE6aoaYb12rxkAsWZqcgMt5aJ0oAg',
    '__cf_bm': '7nvNcd6OzwhKPrJTCCYpxff6K38T9qVn.W8A22xqBao-1790271701.2148075-1.0.1.1-2a61_Em2tr_Y8oThB6HnkalN2h6h9C7IqLb7Jtxz2t7zKoo7weLJlzETq3O3AOL1usXcTxpj38rOiVdfBuobs9aXA6aGgIwo1QHpta2m3OgJ5VPG18ArfuHB5WYcBYnw',
    'forterToken': '90f94d93e1ff4dc2b3f13542a6ef4537_1790271700143_1254_UDF43-m4_23ck_jBRUGURUTcE%3D-11-v2',
    'forterToken': '90f94d93e1ff4dc2b3f13542a6ef4537_1790271700143_1254_UDF43-m4_23ck_jBRUGURUTcE%3D-11-v2',
    '_cq_s': '4b225054a187a9f0b5fdd3d5c52d7b5d.IyYVQk65mtjcpplF:TTyyaodE0gYut4GhVNdtF2CaHZE10XMonsZ40ZzrvDb2xPDuA2C0jxdRkagleqf9emu4tpF86FQWllUdkIr6Kwdk/qgM6kFjAYVSwAEy0jTEASfw71y7j5/p8sO6o7DMJtFxvJe7/1FXQQJgixiKZ/wJc1s1pc2ZW1VsG1j7Ej5J8SmAvqYqbOww6i0B0Ldy58Kpd4FYCrtsg/g1StaLS4d74BVosq83rT6TZq2wHWdW+DqnLiKBgyGXYV6c4TLGbseaMyvT4zel9jEdEO+VBUSlPMvp9QcWPOJJweIipIpI54A9WVTlGF5K1A4q8TbjOj7+BBRokOkyTLYGRhPWtRYg4+f39R527FT4kx79aMPpa2uO82CO0o4uFtEo2NfDCnxc3yLwId9KAbCCKCcsffwLiA4y4tesC0txZMeq+hb7QznoCL7Eel6ygyoB1qhL+Ny4I/b5/Y9zF4NfqOuh4uRkV8+1al9XKuEVxykDUVOmmXHeKYZgSWg09+UVBJNt2vG1x21QQKtczsC8B0hq7i/Upwju0QdD5TvUdpqwCGgW3LpdO2wtd9hSzq7hC+ev3l/hAqAmmOUHFTj4ufk=:f4IhywOJtnXr6qoNLqvsiw==',
    '_gcl_au': '1.1.1885462140.1790168441.-.-.1790175694.402595479.1790256103.1790271711',
    '_ga_KSM221PNDX': 'GS2.1.s1790268930$o9$g1$t1790271711$j32$l0$h0',
    '_rdt_uuid': '1790168463389.9e4b3354-e567-4693-a868-9321f4b8bc7d',
    '_rdt_em': '79f31ac05e00263441ab916654bfe8fdd45b900cac8b818127c5ecb687942813:79f31ac05e00263441ab916654bfe8fdd45b900cac8b818127c5ecb687942813,79f31ac05e00263441ab916654bfe8fdd45b900cac8b818127c5ecb687942813',
    'tatari-cookie-test': '87073793',
    '_uetsid': 'd5165910b74e11f1858e15214aba5a92',
    '_uetvid': 'd5167e60b74e11f180a4a9b4498b483f',
    'IR_PI': 'e39cc4d0-b74e-11f1-bc83-6591ac20c9d4%7C1790358113645',
    'IR_13634': '1790271713648%7C0%7C1790271713648%7C%7C',
    'AWSALBTGCORS': 'kODqzcHUIXV63spY0kD0+r1Jg4bA02yoSfRoScIqQ6ETwEDTj9Hl5nKvcMuEpBEyya4SUaqJL9XHBGE3ruyTGwGR3ST7F1l4kDLtpnBpheno7OV0bqoZNHrMZN2MS8M5JKeO+7MQ+DRWWvuzCvHn2NMD1AFiXyU7Eg2C1FFbgzAD',
    'ttcsid_CGCUGEBC77UAPU79F02G': '1790268931647::O5uXeu-J9C2SHOwP9H_c.8.1790271720191.1',
    '_cq_session': '8.1790268930255.4cRisRB6h0KnVLTq.1790271720012',
    'ttcsid': '1790268931649::Gpd_ATmua5NBClCvrMVd.8.1790271720191.0::1.2788533.2782844::2788301.16.961.325::2791773.110.0',
    '_upw_id.5831': '4d978257-e9f8-416c-a814-d99dbfe4b0d3.1790168441.7.1790271724.1790266861.00e2df4c-fde1-470f-8be7-266081e045cb.4fa7ae82-a900-430b-9fc9-bec83a385494.1ad41b25-48bb-47b6-a5d0-8274f90e4059.1790268927679.250',
    'AWSALBTG': 'KrwRcm1NtIkC7VXMVm/UqYnN38fWGa18LfCVP3d84S+RwV8NKoPpCx1GGL+CfoOEBaeNS1QzlJihnIVazzgkvIuB8LqL9MOkj5bJpu1Mkb3ufeszsUU8jAz2tYLHzNGnrCIu76VN+P1FBr+MbNx5Zigx13FfPZqKvu4elamfc2Ok',
    'AWSALB': 'eOrL65S/YGNMv0tYcPJ1AdDp2SL+VZsRwgQMcZAIhONhQQc9um8ABpAbk7G3ewLIthLIr+LQI1H43z3x2UNA6YWcCokA3OionAkxJy/2VkQOrD+mWMPIo5SzPR9f',
    'AWSALBCORS': 'eOrL65S/YGNMv0tYcPJ1AdDp2SL+VZsRwgQMcZAIhONhQQc9um8ABpAbk7G3ewLIthLIr+LQI1H43z3x2UNA6YWcCokA3OionAkxJy/2VkQOrD+mWMPIo5SzPR9f',
}
headers = {
    'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
    'accept-language': 'zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'priority': 'u=0, i',
    'sec-ch-ua': '"Microsoft Edge";v="153", "Not_A Brand";v="8", "Chromium";v="153"',
    'sec-ch-ua-arch': '"x86"',
    'sec-ch-ua-bitness': '"64"',
    'sec-ch-ua-full-version': '"153.0.4234.48"',
    'sec-ch-ua-full-version-list': '"Microsoft Edge";v="153.0.4234.48", "Not_A Brand";v="8.0.0.0", "Chromium";v="153.0.8010.53"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-model': '""',
    'sec-ch-ua-platform': '"Windows"',
    'sec-ch-ua-platform-version': '"19.0.0"',
    'sec-fetch-dest': 'document',
    'sec-fetch-mode': 'navigate',
    'sec-fetch-site': 'none',
    'sec-fetch-user': '?1',
    'upgrade-insecure-requests': '1',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0',
    # 'cookie': 'cdeb28a6sb=oauth2v2_int_a9eeb64158476b9c19485d065378f0c8; visitor_id=223.112.31.154.1790168435779000; __cflb=02DiuEXPXZVk436fJfSVuuwDqLqkhavJbxwg9CSZdKfEB; fp_token_7c6a6574-f011-4c9a-abdd-9894a102ccef=e7z7UM0kkn51kR5sbNcqpbxUKszIKupoRHgU7vej3hU=; _cq_duid=1.1790168441.b7inZu6F2tLrsTYQ; spt=4b66fc4d-6fc9-412d-b0f6-3297a0cc0ac1; __ps_r=_; __ps_lu=https://www.upwork.com/; __ps_did=pscrb_832e8a34-0258-4b9f-9512-e248ba0dbea0; __ps_fva=1790168463522; __pdst=44722a29abfa42258dd7e8457bb27601; _fbp=fb.1.1790168465552.734314235541176021; tatari-session-cookie=e135e25e-0891-2436-31e9-28bf86e89a38; _ga=GA1.1.1926171419.1790168464; _tt_enable_cookie=1; _ttp=01M377VFRM11E9JE4JMJJ433BV_.tt.1.1790170742549; recognized=6e3cb5ff0eddd1b2; current_organization_uid=2102754721647967795; company_last_accessed=1020210583; ngIaSidebar=o.0*p.*w.0*e.1; uThemeNull=1; tatari-user-cookie=2102754721647967794; device_view=full; g_state={"i_l":0,"i_ll":1790175734586,"i_b":"fZGgeg3w+R/FWn9wWA6tsXqIwPIFM2GIYzN71ezzF4E","i_e":{"enable_itp_optimization":24},"i_et":1790175734586}; console_user=6e3cb5ff0eddd1b2; enabled_ff=!CI12577UniversalSearch,!Fluid,!MP16400Air3Migration,!OTBnrOn,!SFE1028UmaRecruiterRefineImprovements,!SSINavUser,!i18nGA,CI17409DarkModeUI,CmpLibOn,JPAir3,SFE919UmaRecruiterBasicQt3,SSINavUserBpa,TONB2256Air3Migration,TranscendUIOn,WP658TranscendOn,i18nOn; cookie_prefix=; cookie_domain=.upwork.com; XSRF-TOKEN=2102876886097338876; master_access_token=6fbe682b.oauth2v2_int_4da712c908797fe85ac5de87aaa7dbf1; user_uid=2102754721647967794; oauth2_global_js_token=oauth2v2_int_f11a4e1288497c9fc6bfff7696a549a0; fe54365bsb=oauth2v2_int_78532891cf6ebe551188885ea340e167; _cq_suid=1.1790199897.NFaGEAlDDO54ieYK; IR_gbd=upwork.com; _mkto_trk=id:518-RKL-392&token:_mch-upwork.com-e1c5aa472364fc0101c70fcfe2cab9af; country_code=JP; e53a11d0sb=oauth2v2_int_7a4052cd5516c763d3ab62afbf2d3020; ftr_ncd=6; visitor_gql_token=oauth2v2_int_0c54b3723e80d7713c3afa84a961630f; _cfuvid=VNc_KWhdyaT8X2l49WLCihXlR.QLNAFAneaKN4W.Tv0-1790264966.486338-1.0.1.1-4lWUxg86sudD0YgTDT5xm.8Ga..eVAFhXbQ72G3RLSk; umq=1699; umq_cq=1643; _upw_ses.5831=*; __ps_sr=_; __ps_slu=https://www.upwork.com/nx/search/talent?nav_dir=pop&page=1&__cf_chl_tk=mB8kMyvbbaCRYWWyblNnCx6qRkvssoK4Y6N1WKAUE8Y-1790268907-1.0.1.1-AERFS49ht7XLSUccDPSAEQC0ewjbpyh_ZDdlWbA892Y; profile_vv_gql_token=oauth2v2_int_031046fa0b29fec7bd96adaf2a4b16ba; AttachmentsNuxt_vt=oauth2v2_int_c96df3d897b57f8b9c5b84f07af6d8a0; UniversalSearchNuxt_vt=oauth2v2_int_dc47084c47186ee0161a359d2c6ced30; cf_clearance=DUjRdG4sBPmUOglIbaWtK0a3YSNePla9uTWiUThhJlw-1790271700-1.2.1.1-SkgMHqvDiNd4x6TmMlj.JAx86kEZXTiBG2tI83vGN_JzSuXdLP81ZzPK6qegQas6SL44c.rcY3AltqWbTwFziAT6XGVbF6Crvu3PcsFVboe_a66PSn8BT69NedW9bHW2wXIrHfxk9p8SHZtOjbp9lddv7ZGIBEa7lPsVDMfm8L6L91N_gYwtGS7gnWqk9JOWaB5wz6GWfE9q5WKStuU1Kc0qiDLU5FgT8r9DWRYjPyZLbaeEbWriVzG89hwyeTCO4OsR.x4M.4TNZJYCVpsJ4.aymYex7WXU_XsHRlg3vaAoMdD3e9xgHeee4alKk8wObq3ktBFqOPVOuaSEjPOPA5V1ReWFCXLhHVunInyoEXKwkGifSS2JKY7BpqYSNl.i31E.1kecSx95FROxGkCut9MrQf6VwLVlUWWFbKsgi_WVy..pTvzppZNwbbE6aoaYb12rxkAsWZqcgMt5aJ0oAg; __cf_bm=7nvNcd6OzwhKPrJTCCYpxff6K38T9qVn.W8A22xqBao-1790271701.2148075-1.0.1.1-2a61_Em2tr_Y8oThB6HnkalN2h6h9C7IqLb7Jtxz2t7zKoo7weLJlzETq3O3AOL1usXcTxpj38rOiVdfBuobs9aXA6aGgIwo1QHpta2m3OgJ5VPG18ArfuHB5WYcBYnw; forterToken=90f94d93e1ff4dc2b3f13542a6ef4537_1790271700143_1254_UDF43-m4_23ck_jBRUGURUTcE%3D-11-v2; forterToken=90f94d93e1ff4dc2b3f13542a6ef4537_1790271700143_1254_UDF43-m4_23ck_jBRUGURUTcE%3D-11-v2; _cq_s=4b225054a187a9f0b5fdd3d5c52d7b5d.IyYVQk65mtjcpplF:TTyyaodE0gYut4GhVNdtF2CaHZE10XMonsZ40ZzrvDb2xPDuA2C0jxdRkagleqf9emu4tpF86FQWllUdkIr6Kwdk/qgM6kFjAYVSwAEy0jTEASfw71y7j5/p8sO6o7DMJtFxvJe7/1FXQQJgixiKZ/wJc1s1pc2ZW1VsG1j7Ej5J8SmAvqYqbOww6i0B0Ldy58Kpd4FYCrtsg/g1StaLS4d74BVosq83rT6TZq2wHWdW+DqnLiKBgyGXYV6c4TLGbseaMyvT4zel9jEdEO+VBUSlPMvp9QcWPOJJweIipIpI54A9WVTlGF5K1A4q8TbjOj7+BBRokOkyTLYGRhPWtRYg4+f39R527FT4kx79aMPpa2uO82CO0o4uFtEo2NfDCnxc3yLwId9KAbCCKCcsffwLiA4y4tesC0txZMeq+hb7QznoCL7Eel6ygyoB1qhL+Ny4I/b5/Y9zF4NfqOuh4uRkV8+1al9XKuEVxykDUVOmmXHeKYZgSWg09+UVBJNt2vG1x21QQKtczsC8B0hq7i/Upwju0QdD5TvUdpqwCGgW3LpdO2wtd9hSzq7hC+ev3l/hAqAmmOUHFTj4ufk=:f4IhywOJtnXr6qoNLqvsiw==; _gcl_au=1.1.1885462140.1790168441.-.-.1790175694.402595479.1790256103.1790271711; _ga_KSM221PNDX=GS2.1.s1790268930$o9$g1$t1790271711$j32$l0$h0; _rdt_uuid=1790168463389.9e4b3354-e567-4693-a868-9321f4b8bc7d; _rdt_em=79f31ac05e00263441ab916654bfe8fdd45b900cac8b818127c5ecb687942813:79f31ac05e00263441ab916654bfe8fdd45b900cac8b818127c5ecb687942813,79f31ac05e00263441ab916654bfe8fdd45b900cac8b818127c5ecb687942813; tatari-cookie-test=87073793; _uetsid=d5165910b74e11f1858e15214aba5a92; _uetvid=d5167e60b74e11f180a4a9b4498b483f; IR_PI=e39cc4d0-b74e-11f1-bc83-6591ac20c9d4%7C1790358113645; IR_13634=1790271713648%7C0%7C1790271713648%7C%7C; AWSALBTGCORS=kODqzcHUIXV63spY0kD0+r1Jg4bA02yoSfRoScIqQ6ETwEDTj9Hl5nKvcMuEpBEyya4SUaqJL9XHBGE3ruyTGwGR3ST7F1l4kDLtpnBpheno7OV0bqoZNHrMZN2MS8M5JKeO+7MQ+DRWWvuzCvHn2NMD1AFiXyU7Eg2C1FFbgzAD; ttcsid_CGCUGEBC77UAPU79F02G=1790268931647::O5uXeu-J9C2SHOwP9H_c.8.1790271720191.1; _cq_session=8.1790268930255.4cRisRB6h0KnVLTq.1790271720012; ttcsid=1790268931649::Gpd_ATmua5NBClCvrMVd.8.1790271720191.0::1.2788533.2782844::2788301.16.961.325::2791773.110.0; _upw_id.5831=4d978257-e9f8-416c-a814-d99dbfe4b0d3.1790168441.7.1790271724.1790266861.00e2df4c-fde1-470f-8be7-266081e045cb.4fa7ae82-a900-430b-9fc9-bec83a385494.1ad41b25-48bb-47b6-a5d0-8274f90e4059.1790268927679.250; AWSALBTG=KrwRcm1NtIkC7VXMVm/UqYnN38fWGa18LfCVP3d84S+RwV8NKoPpCx1GGL+CfoOEBaeNS1QzlJihnIVazzgkvIuB8LqL9MOkj5bJpu1Mkb3ufeszsUU8jAz2tYLHzNGnrCIu76VN+P1FBr+MbNx5Zigx13FfPZqKvu4elamfc2Ok; AWSALB=eOrL65S/YGNMv0tYcPJ1AdDp2SL+VZsRwgQMcZAIhONhQQc9um8ABpAbk7G3ewLIthLIr+LQI1H43z3x2UNA6YWcCokA3OionAkxJy/2VkQOrD+mWMPIo5SzPR9f; AWSALBCORS=eOrL65S/YGNMv0tYcPJ1AdDp2SL+VZsRwgQMcZAIhONhQQc9um8ABpAbk7G3ewLIthLIr+LQI1H43z3x2UNA6YWcCokA3OionAkxJy/2VkQOrD+mWMPIo5SzPR9f',
}
def content(pages):
    for page in range(1, pages + 1):
        params = {
        'nav_dir': 'pop',
    } #这里直接修改里面的page参数应该能达到同样的效果吧
        response = requests.get(f'https://www.upwork.com/nx/search/talent?nav_dir=pop&page={page}', params=params, cookies=cookies, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        html = etree.HTML(response.text)

        names = soup.find_all('h5', class_='name m-0 rr-mask') #姓名
        for name in names:
            name = name.text.replace('.','')
            a.append(name)

        unknows = soup.find_all('h4', class_='title m-0') #自我介绍？ and 界面跳转
        # for unk in unknows:
        #     href = unk.find_all('a')              # 跳转后的信息爬取不知到为什么老失效 故待优化
        #     href = href[0].get('href')
        #     # href = re.sub(r'href="(\s\S+?)"',r'\g<1>', href)
        #     href = 'https://www.upwork.com' + href
        #     i.append(href)

        for unk in unknows:
            unk = unk.text
            b.append(unk)
        contrys = soup.find_all('p', class_='m-0 location rr-mask') #来自哪里 国家
        for contry in contrys:
            c.append(contry.text)
        contents = soup.find_all('div', class_='details mt-3x')
        for content in contents:
        #渴望薪资

            pay = re.findall(r'<span data-test="rate-per-hour">[\s\S]+?</span>', str(content))[0]
            pay = re.sub(r'<span data-test="rate-per-hour">([\s\S]+?)</span>',r'\g<1>', pay)
            pay = pay + '/hr'
            d.append(pay)

        #工作成功率
            try:
                success_rate = re.findall(r'--><span>[\s\S]+?%</span><!--',str(content))[0]
                success_rate = re.sub(r'--><span>([\s\S]+?%)</span><!--',r'\g<1>',success_rate)
                e.append(success_rate)
            except:
                e.append('没有')

        #赚了多少
            try:
                earned = content.find_all('div', class_='details-item text-base rr-mask')[2]
                earned = re.findall(r'<strong>[\s\S]+?</strong>', str(earned))[0]
                earned = re.sub(r'<strong>([\s\S]+?)</strong>', r'\g<1>', earned).replace(' earned','')
                f.append(earned)
            except:
                f.append('没有')

        #自我介绍
        intros = soup.find_all('div', class_='air3-line-clamp-wrapper description text-body mt-3x')
        for intro in intros:
            intro = re.findall(r'-->[\s\S]+?<!--', str(intro))[1].replace('-->', '').replace('<!--', '')
            g.append(intro)
        #技能
        skills = soup.find_all('div', class_='mt-3x')
        for skill in skills:
            skill = skill.find_all('div', class_='air3-token-container')
            if len(skill) == 0:
                continue
        # print(skill)
            for i in skill:
                button = re.findall(r'<button class="air3-token" data-v-63bb444d="" type="button">[\s\S]+?/button>', str(i))[0]
                button = button.replace('<button class="air3-token" data-v-63bb444d="" type="button">','').replace('</button>','')
            # button = re.findall(r'<button class="air3-token" data-v-63bb444d="" type="button">[\s\S]+?</button>', str(button[0]))
                h.append(button)
        #由于跳转后爬取失败，所爬取的内容较少
    seen = set() #去重
    for aa, bb, cc, dd ,ee, ff, gg, hh in zip(a, b, c, d, e, f, g, h):
            bb_clean = bb.strip() if aa else ""
            if bb_clean and bb_clean not in seen:
                seen.add(bb_clean)
                sheet.append([aa,bb,cc,dd,ee,ff,gg,hh])
    wb.save('求职人员.xlsx')
    # -------------------------------------------------------------------照片部分
def photos():
    for page in range(1, pages + 1):
        params = {
        'nav_dir': 'pop',
    }
        response = requests.get(f'https://www.upwork.com/nx/search/talent?nav_dir=pop&page={page}', params=params,
                                cookies=cookies, headers=headers)
        soup = BeautifulSoup(response.text, 'html.parser')
        imgs = soup.find_all('img', class_='air3-avatar air3-avatar-60')  # 照骗
        for img in imgs: # 不知道为什么吗 爬取多页时会缺
            url = img['src']
            response = requests.get(url, cookies=cookies, headers=headers)
            path = 'imgs'
            if not os.path.exists(path):
                os.makedirs(path)
            with open(os.path.join(path, img['alt'] + '.jpg'), 'wb') as f:
                f.write(response.content)

# def demo():
#     for page in range(1, pages + 1):
#         params = {
#         'nav_dir': 'pop',
#     }
#     response = requests.get(f'https://www.upwork.com/nx/search/talent?nav_dir=pop&page={page}', params=params, cookies=cookies, headers=headers)
#     soup = BeautifulSoup(response.text, 'html.parser')
#     unknows = soup.find_all('h4', class_='title m-0')
#     for unk in unknows:
#         href = unk.find_all('a')              # 跳转后的信息爬取不知到为什么老失效 故待优化
#         href = href[0].get('href')
#         # href = re.sub(r'href="(\s\S+?)"',r'\g<1>', href)
        # href = 'https://www.upwork.com' + href
        # res = requests.get(href, params=params,cookies=cookies, headers=headers)
        # soups = BeautifulSoup(res.text, 'html.parser')
        # print(soups)
        # print('-------------------------------------')


if __name__ == '__main__':
    pass
    photos()
    content(pages)
    # demo()
