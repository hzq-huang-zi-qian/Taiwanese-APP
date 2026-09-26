#!/usr/bin/env python3
"""Build EP1 second-half Taiwanese database Excel with even time distribution (19:00–37:00)."""

import importlib.util
import json
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
TAIGI_SRT = BASE / "_MConverter.eu_《夜市人生》第一集後半段完整190句台語漢字字幕檔 (1).srt"
OUTPUT = BASE / "台語資料庫_EP1後半.xlsx"

TIME_START_SEC = 19 * 60       # 19:00
TIME_END_SEC = 37 * 60         # 37:00
TARGET_COUNT = 50
MAX_CUE_DURATION = 25.0

SKIP_EXACT = {
    "好", "有啊", "好喔", "是", "是喔", "爸", "我", "對...", "對無？",
    "月霞", "慶祥", "友慧", "友志", "漢良", "娜娜", "如意", "爸爸", "頭家",
    "來…", "來...", "算帳？", "無無無無", "好啊！", "你！", "啊！", "歹勢。",
    "不知影？", "做一個。", "對啦", "是", "無按呢啦",
}

# Load shared tagging / labels from first-half builder
_spec = importlib.util.spec_from_file_location(
    "build_ep1", BASE / "build_taigi_database.py"
)
_build = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_build)

LT_LABELS = _build.LT_LABELS
SYNTAX_IDS = _build.SYNTAX_IDS
SEMANTICS_IDS = _build.SEMANTICS_IDS
PHONOLOGY_IDS = _build.PHONOLOGY_IDS
PRAGMATICS_IDS = _build.PRAGMATICS_IDS
TAG_EXPLANATIONS = _build.TAG_EXPLANATIONS
infer_linguistic_tags = _build.infer_linguistic_tags
normalize_text = _build.normalize_text
categorize_tags = _build.categorize_tags
build_explanation = _build.build_explanation
ensure_pragmatics = _build.ensure_pragmatics

PINYIN_OVERRIDES = {
    **getattr(_build, "PINYIN_OVERRIDES", {}),
    "暗時生意好無？": "Àm-sî seng-ì hó--bô?",
    "你轉來啦！": "Lí tńg-lâi--lah!",
    "你轉來了啊！": "Lí tńg-lâi--ah!",
    "你有乖無？": "Lí ū kuai--bô?",
    "你公司的人工發予人未？": "Lí kong-si ê jîn-kang hoat hōo lâng buē?",
    "債務還清未？": "Sāi-bu̍t huan-tshing buē?",
    "還清矣。": "Huan-tshing--ah.",
    "我看你乾脆共公司收收咧。": "Guá khuànn lí tshian-tshui kā kong-si siu-siu--leh.",
    "猶未食咧，腹肚真枵。": "Iáu-bue̍h tsia̍h--leh, pa̍k-tó͘ tsin kiau.",
    "慢且爸爸一爿食、一爿陪你做功課，好無？": "Bān-chhiáⁿ pa-pa tsi̍t-pêng tsia̍h, tsi̍t-pêng phuè lí tsò-kong-khòo, hó--bô?",
    "友慧，爸爸共你講。": "Iú-huì, pa-pa kā lí kóng.",
    "這學期愛認真一滴仔。": "Tsit ha̍k-kî ài jīn-tsin tsi̍t-tih-á.",
    "假使你會得考第一名。": "Ká-sú lí ē-tit khó tē-it-miâ.",
    "爸爸是不是共你拼矣？": "Pa-pa sī-m̄-sī kā lí piànn--ah?",
    "你也同款。": "Lí iā tông-khuán.",
    "來，趕緊咧。": "Lâi, kín-kín--leh.",
    "歹勢，這張無人坐，借一下。": "Pháiⁿ-sè, tsit tiuⁿ bô-lâng tsē, chai tsi̍t-ē.",
    "留予你坐的。": "Lâu hōo lí tsē ê.",
    "徛咧！": "Khiā--leh!",
    "你猶想欲覕無？": "Lí iáu siūnn-beh bē--bô?",
    "借問你是？": "Tshiò-mn̄g lí sī?",
    "我來揣李慶祥算帳。": "Guá lâi tshuē Lí Khìng-siông sǹg-tiūnn.",
    "我翁欠你錢無？": "Guá ang khiàm lí tsînn bô?",
    "對啦對啦，欠伊一滴滴仔，我會共伊解決。": "Tùi--lah tùi--lah, khiàm i tsi̍t-tih-tih-á, guá ē kā i kái-kuat.",
    "你無錢是要共人解決啥？": "Lí bô tsînn sī beh kā lâng kái-kuat siáⁿ?",
    "借問我翁欠你偌濟錢？": "Tshiò-mn̄g guá ang khiàm lí gōa-tsuē tsînn?",
    "李慶祥欠我的。": "Lí Khìng-siông khiàm guá ê.",
    "毋單單是錢。": "M̄ tan-tan sī tsînn.",
    "伊拄去大陸的時。": "I tú khì Tāi-lio̍k ê sî.",
    "人生地不熟，靠我熟似客群。": "Jîn-sinn tē put-se̍k, khò guá se̍k-sāi kheh-kûn.",
    "我未曾在伊身上提著任何好處。": "Guá buē-tsài tī i sin-siōng the̍h-tio̍h jīm-hô hó-chù.",
    "我提錢予伊周轉。": "Guá the̍h-tsînn hōo i tsiu-tsún.",
    "伊是你共伊生的囡仔？": "I sī lí kā i senn ê gín-á?",
    "你白賊！": "Lí pe̍h-chha̍t!",
    "我在大陸這款拚死拚活。": "Guá tī Tāi-lio̍k tsit-khuán piànn-sí piànn-ua̍h.",
    "也是欲拚予恁過較好日。": "Iā-sī beh piànn hōo lín kuè khah hó ji̍t.",
    "我就運氣歹。": "Guá tō ūn-khì phái.",
    "提去博股票、炒期貨。": "The̍h-khì po̍k kóo-phiò, tshá kî-huò.",
    "賺的錢攏開了了。": "Thàn ê tsînn lóng khai liáu-liáu.",
    "才著靠娜娜共我周轉替我還債務。": "Tsiah tio̍h khò Ná-na kā guá tsiu-tsún thè guá huan sāi-bu̍t.",
    "著是因為按呢，咱兩個才會做伙。": "Tio̍h sī in-uī án-ne, lán nn̄g ê tsiah ē tsò-hué.",
    "蛤？": "Hah?",
    "你為了欲予妻兒過好日。": "Lí ūi-liáu beh hōo chhài-jî kuè hó ji̍t.",
    "著騙我的錢、我的人無？": "Tio̍h phian guá ê tsînn, guá ê lâng bô?",
    "阮兩個做伙是順其自然。": "Guán nn̄g ê tsò-hué sī sūn-kî tsì-jiân.",
    "我哪有騙你。": "Guá ná-ū phian lí.",
    "啥物順其自然啦！": "Siáⁿ-mih sūn-kî tsì-jiân--lah!",
    "你歸句話講甲這順、這自然。": "Lí kui-kù uē kóng kah tsit sūn, tsit tsì-jiân.",
    "你可以為了錢，連自家的妻兒攏無愛。": "Lí ē-tit uī-liáu tsînn, liân ka-kī ê chhài-jî lóng bô ài.",
    "你順其自然共別的查某做伙。": "Lí sūn-kî tsì-jiân kā pa̍t ê cha-bó͘ tsò-hué.",
    "做男人按怎可能為了錢，無愛自己的妻兒？": "Tsò lâm-lâng án-nuá khó-lêng uī-liáu tsînn, bô ài ka-kī ê chhài-jî?",
    "我講欲錢會得啦，轉去共蔡月霞離婚。": "Guá kóng beh tsînn ē-tit--lah, tńg-khì kā Tshài Gua̍t-hā lî-hun.",
    "這件事你共我拖真久矣喔。": "Tsit kiāⁿ sū lí kā guá thoa tsin kú--ah--ooh.",
    "雖然伊逼我共你離婚。": "Suî-jiân i pik guá kā lí lî-hun.",
    "但是我猶原無共你離婚。": "Tān-sī guá iáu-gôan bô kā lí lî-hun.",
    "我只是轉來共你提錢而已啊。": "Guá tsi̍s tńg-lâi kā lí the̍h-tsînn tú-ī--ah.",
    "按呢真有志氣，是無？": "Án-ne tsin ū chì-khì, sī--bô?",
    "好，你講袂拋妻棄子。": "Hó, lí kóng bē pha-chhài khì-kiáⁿ.",
    "這馬我幫你還錢矣。": "Tsit-má guá pāng lí huan-tsînn--ah.",
    "這馬，隨共伊一刀兩斷！": "Tsit-má, suî kā i tsi̍t-to tnn̄g tn̄g!",
    "袂使！": "Bē-sái!",
    "我今仔日是來𤆬人的，慶祥今仔日愛共我行。": "Guá kin-á-ji̍t sī lâi tshuā lâng ê, Khìng-siông kin-á-ji̍t ài kā guá kiânn.",
    "你憑啥物？": "Lí pîng siáⁿ-mih?",
    "憑伊欺騙我。": "Pîng i khi-phian guá.",
    "你敢拍我？": "Lí kám phah guá?",
    "我拍你又閣安怎？": "Guá phah lí iū koh án-nuá?",
    "攏是我自個兒賺的。": "Lóng sī guá tsū-kè-jī thàn ê.",
    "好矣，娜娜。": "Hó--ah, Ná-na.",
    "是你翁無愛你。": "Sī lí ang bô ài lí.",
    "好矣好矣...莫按呢。": "Hó--ah hó--ah... mài án-ne.",
    "你這個痟查某！": "Lí tsit ê siáu cha-bó͘!",
    "你這個查某太超過矣！搶人翁搶到厝內來！大家鬥相共月霞姊，共這隻狐狸精拖出去！": "Lí tsit ê cha-bó͘ thài tshiau-kòe--ah! Chhiúⁿ lâng-ang chhiúⁿ kàu tshù-lāi lâi! Ta-ke tàu-saⁿ-kāng Gua̍t-hā tsí, kā tsit chiah hô͘-lî-tsing thoa--tshut-khì!",
    "你無快行，拍甲你現出原形！": "Lí bô kín kiânn, phah kah lí hiàn-tshut guân-hêng!",
    "假使你敢一人對眾人。": "Ká-sú lí kám tsi̍t-lâng tuì tsiòng-jîn.",
    "阮吐喙涎著渰死你矣。": "Guán thò͘ tshuì-nn̄g tio̍h lam-sí lí--ah.",
    "拜託好無？": "Pài-thok hó--bô?",
    "阮兜都已經一團亂矣。": "Guán tau to í-king tsi̍t-thoân luān--ah.",
    "恁莫佇這湊熱鬧矣。": "Lín mài tī tsia tshàu jia̍t-lāu--ah.",
    "寡不敵眾。": "Kóa put-tik tsiòng.",
    "放落杓仔，甲無代誌。": "Pàng-lo̍h giâu-á, kah bô tāi-tsì.",
    "這馬我愛討轉來！": "Tsit-má guá ài thó--tńg-lâi!",
    "就是按呢。": "Tō-sī án-ne.",
    "才著靠娜娜共我周轉。": "Tsiah tio̍h khò Ná-na kā guá tsiu-tsún.",
    "會予地下錢莊收去矣。": "Ē hōo tē-hā tsînn-tsng siu--khì--ah.",
    "按呢你打算欲拋妻棄子著對啦？": "Án-ne lí tásǹg beh pha-chhài khì-kiáⁿ tio̍h tùi--lah?",
    "你予人追拍討債的時。": "Lí hōo lâng tui phah thó-tsài ê sî.",
    "是我擋佇頭前。": "Sī guá tòng tī thâu-chêng.",
    "我四界拜託。": "Guá sì-kài pài-thok.",
    "煩惱你的代誌煩惱到睏袂去。": "Huân-ló lí ê tāi-tsì huân-ló kàu khùn bē--khì.",
    "換來是這款的結果無？": "Hōan--lâi sī tsit khuán ê kiat-kó bô?",
    "慶祥，共我上車。": "Khìng-siông, kā guá siōng-tshia.",
    "是你向阮保證，你翁三天內會提錢予阮。": "Sī lí hiòng guán pó-tsìng, lí ang sann-kang lāi ē the̍h-tsînn hōo guán.",
    "這馬，伊人溜走矣。": "Tsit-má, i lâng liû--ah.",
    "阮當然揣你討！": "Guán tong-jiân tshuē lí thó!",
    "我不知影慶祥無還恁錢。": "Guá m̄ tsai-iánn Khìng-siông bô huan lín tsînn.",
    "你一聲不知影著無代誌矣無？": "Lí tsi̍t-siann m̄-tsai-iánn tio̍h bô tāi-tsì--ah bô?",
    "你看，我共如意。": "Lí khuànn, guá kā Jû-ì.",
    "月霞，總而言之。": "Gua̍t-hā, tsóng-jî-yan-chi.",
    "今仔日你若無共我一個交代。": "Kin-á-ji̍t lí nā bô kā guá tsi̍t ê kau-tài.",
    "你著莫怨嘆阮十幾年的鄰居。": "Lí tio̍h mài uàn-thàn guán tsa̍p-kuí nî ê lîn-ki.",
    "逼急矣我揣大家去法院告你。": "Pik-kip--ah guá tshuē ta-ke khì hoat-īnn kòo lí.",
    "漢良，大家遮多年的鄰居矣。": "Hàn-liông, ta-ke tsia to-nî ê lîn-ki--ah.",
    "何必安呢，有代誌慢慢講。": "Hô-pit án-ne, ū tāi-tsì bān-bān kóng.",
    "黃昏紅紅的夕陽，就是咱的日出！": "Hông-hun âng-âng ê se̍k-iông, tō-sī lán ê ji̍t-tshut!",
    "咱是在夜市討生活的人，天色愈暗咱愈愛拍拚！": "Lán sī tī iā-tshī thó-sing-ua̍h ê lâng, thian-se̍k jú àm lán jú ài phah-piànn!",
    "人溜得一乾二淨。": "Lâng liû-tit tsi̍t-kan-jī-tsīng.",
    "我提予伊這寡錢。": "Guá the̍h hōo i tsit kó tsînn.",
    "我共伊拍大哥大攏袂通，無接。": "Guá kā i phah tāi-ko-tāi lóng bē-thang, bô tsih.",
    "著親像斷線的風吹。": "Tio̍h chhin-chhiūⁿ tn̄g-sòaⁿ ê hong-tshue.",
    "一去不越頭。": "It-khì put-jia̍t-thâu.",
    "細錢，忍一下著過去矣。": "Sè-tsînn, jím tsi̍t-ē tio̍h kuè-khì--ah.",
    "有志，拜託共我提一桶水。": "Iú-chì, pài-thok kā guá the̍h tsi̍t tháng-tsuí.",
    "閣無中！": "Koh bô tiong!",
    "你十個攏無擲中，對無？": "Lí tsa̍p ê lóng bô tàn-tiong, tùi--bô?",
    "阮照約定行事，對無？": "Guán tsiàu iok-tēng hêng-sū, tùi--bô?",
    "外插兩百塊，拿來！": "Guā-tshap nn̄g-pah khoài, ná--lâi!",
    "無啦，毋是...毋是按呢啦。": "Bô--lah, m̄-sī... m̄-sī án-ne--lah.",
    "無共我講伊有家庭，著共我生這個囡仔，伊著對我共這個囡仔負責任！": "Bô kā guá kóng i ū ka-tîng, tio̍h kā guá senn tsit ê gín-á, i tio̍h tuì guá kā tsit ê gín-á hū-tsek-jīm!",
    "阮母仔囝開的錢。": "Guán bú-á-kiáⁿ khai ê tsînn.",
    "我無睏無歇忙甲按呢。": "Guá bô khùn bô hioh bông kah án-ne.",
    "友志、友慧猶會共我做生意。": "Iú-chì, Iú-huì iáu ē kā guá tsò-sing-ì.",
    "告上法院欲創啥？": "Kòo siōng hoat-īnn beh tshòng siáⁿ?",
    "看你敢會得商量啊。": "Khuànn lí kám ē-tit siong-lióng--ah.",
    "等一咧喔！": "Tán tsi̍t--leh--ooh!",
    "共我周轉替我還債務。": "Kā guá tsiu-tsún thè guá huan sāi-bu̍t.",
    "也不知影伊提去佗位。": "Iā m̄ tsai-iánn i the̍h-khì tó-uī.",
}

MANDARIN_OVERRIDES = {
    "暗時生意好無？": "晚上生意好不好？",
    "你轉來啦！": "你回來啦！",
    "你轉來了啊！": "你回來了啊！",
    "你有乖無？": "你有乖嗎？",
    "你公司的人工發予人未？": "你公司的工資發給人了嗎？",
    "債務還清未？": "債務還清了嗎？",
    "還清矣。": "還清了。",
    "我看你乾脆共公司收收咧。": "我看你乾脆把公司收一收吧。",
    "猶未食咧，腹肚真枵。": "還沒吃呢，肚子真餓。",
    "慢且爸爸一爿食、一爿陪你做功課，好無？": "等一下爸爸一邊吃、一邊陪你寫功課，好嗎？",
    "友慧，爸爸共你講。": "友慧，爸爸跟你說。",
    "這學期愛認真一滴仔。": "這學期要認真一點。",
    "假使你會得考第一名。": "假設你能考第一名。",
    "爸爸是不是共你拼矣？": "爸爸是不是為你打拼了？",
    "你也同款。": "你也一樣。",
    "來，趕緊咧。": "來，趕快。",
    "歹勢，這張無人坐，借一下。": "不好意思，這張沒人坐，借一下。",
    "留予你坐的。": "留給你坐的。",
    "徛咧！": "站起來！",
    "你猶想欲覕無？": "你還想裝沒看到嗎？",
    "借問你是？": "請問你是？",
    "我來揣李慶祥算帳。": "我來找李慶祥算帳。",
    "我翁欠你錢無？": "我老公欠你錢嗎？",
    "對啦對啦，欠伊一滴滴仔，我會共伊解決。": "對啦對啦，欠他一點點，我會幫他解決。",
    "你無錢是要共人解決啥？": "你沒錢是要幫人解決什麼？",
    "借問我翁欠你偌濟錢？": "請問我老公欠你多少錢？",
    "李慶祥欠我的。": "李慶祥欠我的。",
    "毋單單是錢。": "不只是錢。",
    "伊拄去大陸的時。": "他剛去大陸的時候。",
    "人生地不熟，靠我熟似客群。": "人生地不熟，靠我熟悉的客群。",
    "我未曾在伊身上提著任何好處。": "我從未在他身上得到任何好處。",
    "我提錢予伊周轉。": "我拿錢給他周轉。",
    "伊是你共伊生的囡仔？": "他是你跟他生的孩子嗎？",
    "你白賊！": "你騙人！",
    "我在大陸這款拚死拚活。": "我在大陸這樣拚死拚活。",
    "也是欲拚予恁過較好日。": "也是要拚給你們過更好的日子。",
    "我就運氣歹。": "我就運氣壞。",
    "提去博股票、炒期貨。": "拿去賭股票、炒期貨。",
    "賺的錢攏開了了。": "賺的錢都賠光了。",
    "才著靠娜娜共我周轉替我還債務。": "才要靠娜娜周轉幫我還債務。",
    "著是因為按呢，咱兩個才會做伙。": "就是因為這樣，我們兩個才會在一起。",
    "蛤？": "蛤？",
    "你為了欲予妻兒過好日。": "你為了要讓妻兒過好日子。",
    "著騙我的錢、我的人無？": "要騙我的錢、我的人嗎？",
    "阮兩個做伙是順其自然。": "我們兩個在一起是順其自然。",
    "我哪有騙你。": "我哪有騙你。",
    "啥物順其自然啦！": "什麼順其自然啦！",
    "你歸句話講甲這順、這自然。": "你這句話說得這麼順、這麼自然。",
    "你可以為了錢，連自家的妻兒攏無愛。": "你可以為了錢，連自己的妻兒都不愛。",
    "你順其自然共別的查某做伙。": "你順其自然跟別的女人在一起。",
    "做男人按怎可能為了錢，無愛自己的妻兒？": "做男人怎麼可能為了錢，不愛自己的妻兒？",
    "我講欲錢會得啦，轉去共蔡月霞離婚。": "我說要錢可以，回去跟蔡月霞離婚。",
    "這件事你共我拖真久矣喔。": "這件事你跟我拖很久了喔。",
    "雖然伊逼我共你離婚。": "雖然她逼我跟你離婚。",
    "但是我猶原無共你離婚。": "但是我還是沒有跟你離婚。",
    "我只是轉來共你提錢而已啊。": "我只是回來跟你拿錢而已啊。",
    "按呢真有志氣，是無？": "這樣真有骨氣，是不是？",
    "好，你講袂拋妻棄子。": "好，你說不會拋妻棄子。",
    "這馬我幫你還錢矣。": "現在我幫你還錢了。",
    "這馬，隨共伊一刀兩斷！": "現在，馬上跟他一刀兩斷！",
    "袂使！": "不行！",
    "我今仔日是來𤆬人的，慶祥今仔日愛共我行。": "我今天來帶人的，慶祥今天要跟我走。",
    "你憑啥物？": "你憑什麼？",
    "憑伊欺騙我。": "憑他欺騙我。",
    "你敢拍我？": "你敢打我？",
    "我拍你又閣安怎？": "我打你又怎樣？",
    "攏是我自個兒賺的。": "都是我自個兒賺的。",
    "好矣，娜娜。": "好了，娜娜。",
    "是你翁無愛你。": "是你老公不愛你。",
    "好矣好矣...莫按呢。": "好了好了……別這樣。",
    "你這個痟查某！": "你這個瘋女人！",
    "你這個查某太超過矣！搶人翁搶到厝內來！大家鬥相共月霞姊，共這隻狐狸精拖出去！": "你這個女人太過分了！搶人家老公搶到家裡來！大家幫月霞姐，把這狐狸精拖出去！",
    "你無快行，拍甲你現出原形！": "你不快走，打得你現出原形！",
    "假使你敢一人對眾人。": "假設你敢一個人對众人。",
    "阮吐喙涎著渰死你矣。": "我們吐口水都能淹死你。",
    "拜託好無？": "拜託好不好？",
    "阮兜都已經一團亂矣。": "我們家都已經一團亂了。",
    "恁莫佇這湊熱鬧矣。": "你們別在這湊熱鬧了。",
    "寡不敵眾。": "寡不敵眾。",
    "放落杓仔，甲無代誌。": "放下勺子，沒事。",
    "這馬我愛討轉來！": "現在我要討回來！",
    "就是按呢。": "就是這樣。",
    "才著靠娜娜共我周轉。": "才要靠娜娜周轉我。",
    "會予地下錢莊收去矣。": "會被地下錢莊收走了。",
    "按呢你打算欲拋妻棄子著對啦？": "這樣你打算拋妻棄子是對的啦？",
    "你予人追拍討債的時。": "你被人追打討債的時候。",
    "是我擋佇頭前。": "是我擋在前面。",
    "我四界拜託。": "我四處拜託。",
    "煩惱你的代誌煩惱到睏袂去。": "煩惱你的事煩惱到睡不著。",
    "換來是這款的結果無？": "換來是這種結果嗎？",
    "慶祥，共我上車。": "慶祥，扶我上車。",
    "是你向阮保證，你翁三天內會提錢予阮。": "是你向我們保證，你老公三天內會拿錢給我們。",
    "這馬，伊人溜走矣。": "現在，他們溜走了。",
    "阮當然揣你討！": "我們當然找你討！",
    "我不知影慶祥無還恁錢。": "我不知道慶祥沒還你們錢。",
    "你一聲不知影著無代誌矣無？": "你一句不知道就沒事了嗎？",
    "你看，我共如意。": "你看，我對如意。",
    "月霞，總而言之。": "月霞，總而言之。",
    "今仔日你若無共我一個交代。": "今天你若不給我一個交代。",
    "你著莫怨嘆阮十幾年的鄰居。": "你別埋怨我們十幾年的鄰居情分。",
    "逼急矣我揣大家去法院告你。": "逼急了我就找大家去法院告你。",
    "漢良，大家遮多年的鄰居矣。": "漢良，大家這裡多年的鄰居了。",
    "何必安呢，有代誌慢慢講。": "何必這樣，有事慢慢說。",
    "黃昏紅紅的夕陽，就是咱的日出！": "黃昏紅紅的夕陽，就是我們的日出！",
    "咱是在夜市討生活的人，天色愈暗咱愈愛拍拚！": "我們是在夜市討生活的人，天越黑我們越愛打拼！",
    "人溜得一乾二淨。": "人溜得一乾二淨。",
    "我提予伊這寡錢。": "我拿給他這些錢。",
    "我共伊拍大哥大攏袂通，無接。": "我給他打手機都打不通，沒接。",
    "著親像斷線的風吹。": "就像斷線的風箏。",
    "一去不越頭。": "一去不回頭。",
    "細錢，忍一下著過去矣。": "小錢，忍一下就能過去了。",
    "有志，拜託共我提一桶水。": "有志，拜託幫我提一桶水。",
    "閣無中！": "又沒中！",
    "你十個攏無擲中，對無？": "你十個都沒投中，對不對？",
    "阮照約定行事，對無？": "我們照約定行事，對不對？",
    "外插兩百塊，拿來！": "外加兩百塊，拿來！",
    "無啦，毋是...毋是按呢啦。": "沒有啦，不是……不是這樣啦。",
    "無共我講伊有家庭，著共我生這個囡仔，伊著對我共這個囡仔負責任！": "沒跟我說他有家庭，還跟我生了這個孩子，他應該對我跟這個孩子負責任！",
    "阮母仔囝開的錢。": "我們母子開的錢。",
    "我無睏無歇忙甲按呢。": "我沒睡沒休息忙成這樣。",
    "友志、友慧猶會共我做生意。": "友志、友慧還會跟我做生意。",
    "告上法院欲創啥？": "告上法院想幹嘛？",
    "看你敢會得商量啊。": "看你能不能商量啊。",
    "等一咧喔！": "等一下喔！",
    "共我周轉替我還債務。": "幫我周轉替我還債務。",
    "也不知影伊提去佗位。": "也不知道他拿到哪裡去了。",
}


def fix_time_line(line: str) -> str:
    line = line.strip()
    line = re.sub(r"(\d{2}:\d{2})\s+(\d{2},)", r"\1:\2", line)
    line = re.sub(r"\s+", "", line)
    line = line.replace(".", ",")
    return line


def parse_srt(path: Path) -> list[dict]:
    raw = path.read_text(encoding="utf-8")
    raw = re.sub(r"(\d{2}:\d{2})\s+(\d{2},)", r"\1:\2", raw)
    lines = raw.splitlines()
    entries = []
    i = 0
    while i < len(lines):
        while i < len(lines) and not lines[i].strip():
            i += 1
        if i >= len(lines):
            break
        if re.match(r"^\d+$", lines[i].strip()):
            i += 1
        if i >= len(lines):
            break
        time_line = fix_time_line(lines[i])
        i += 1
        if "-->" not in time_line:
            continue
        text_lines = []
        while i < len(lines):
            s = lines[i].strip()
            if not s:
                i += 1
                break
            nxt = fix_time_line(lines[i + 1]) if i + 1 < len(lines) else ""
            if re.match(r"^\d+$", s) and "-->" in nxt:
                break
            text_lines.append(s)
            i += 1
        m = re.search(
            r"(\d{2}):(\d{2}):(\d{2}),(\d{3})-->(\d{2}):(\d{2}):(\d{2}),(\d{3})",
            time_line,
        )
        if not m or not text_lines:
            continue
        sh, sm, ss, sms, eh, em, es, ems = map(int, m.groups())
        start = sh * 3600 + sm * 60 + ss + sms / 1000
        end = eh * 3600 + em * 60 + es + ems / 1000
        if end < start:
            end = start + 1.5
        if end - start > MAX_CUE_DURATION:
            if end >= TIME_START_SEC:
                start = max(end - 3.0, TIME_START_SEC)
            else:
                continue
        text = " ".join(text_lines).strip()
        if not text:
            continue
        entries.append({
            "text": text,
            "start": start,
            "end": end,
            "start_ms": int(start * 1000),
            "timestamp": time_line,
            "timestamp_range": (
                f"{int(start//3600):02d}:{int((start%3600)//60):02d}:{int(start%60):02d}-"
                f"{int(end//3600):02d}:{int((end%3600)//60):02d}:{int(end%60):02d}"
            ),
        })
    return entries


def is_skip(text: str) -> bool:
    t = text.strip()
    if t in SKIP_EXACT:
        return True
    if len(t) < 4:
        return True
    if re.fullmatch(r"[A-Za-z\u4e00-\u9fff]{1,3}[！!？?。…]*", t):
        return True
    if re.fullmatch(r"(對|無|莫|好)[啦啊喔…\.!！?？]*", t):
        return True
    return False


def lookup_pinyin(q: str) -> str:
    if q in PINYIN_OVERRIDES:
        return PINYIN_OVERRIDES[q]
    nq = normalize_text(q)
    for k, v in PINYIN_OVERRIDES.items():
        if normalize_text(k) == nq:
            return v
    return ""


def lookup_mandarin(q: str) -> str:
    if q in MANDARIN_OVERRIDES:
        return MANDARIN_OVERRIDES[q]
    nq = normalize_text(q)
    for k, v in MANDARIN_OVERRIDES.items():
        if normalize_text(k) == nq:
            return v
    return ""


def score_candidate(c: dict) -> tuple:
    return (len(c["tags"]), len(c["taigi"]), -c["start_ms"])


def select_evenly(candidates: list[dict], target: int) -> list[dict]:
    """Pick one tagged sentence per time bin across 19:00–37:00."""
    bin_w = (TIME_END_SEC - TIME_START_SEC) / target
    used = set()
    selected = []

    for i in range(target):
        lo = TIME_START_SEC + i * bin_w
        hi = TIME_START_SEC + (i + 1) * bin_w
        center = (lo + hi) / 2

        pool = [c for c in candidates if normalize_text(c["taigi"]) not in used]
        in_bin = [c for c in pool if lo <= c["start_ms"] / 1000 < hi]

        if not in_bin:
            in_bin = sorted(
                pool,
                key=lambda c: abs(c["start_ms"] / 1000 - center),
            )[:8]

        if not in_bin:
            continue

        pick = max(in_bin, key=score_candidate)
        selected.append(pick)
        used.add(normalize_text(pick["taigi"]))

    return sorted(selected, key=lambda c: c["start_ms"])


def main():
    entries = parse_srt(TAIGI_SRT)
    candidates = []

    for row in entries:
        q = row["text"]
        sec = row["start"]
        if sec < TIME_START_SEC or sec > TIME_END_SEC:
            continue
        if is_skip(q):
            continue

        pinyin = lookup_pinyin(q)
        tags = infer_linguistic_tags(q, pinyin)
        if not tags:
            continue
        if pinyin:
            tags = sorted(set(tags) | set(infer_linguistic_tags(q, pinyin)))

        candidates.append({
            "taigi": q,
            "pinyin": pinyin,
            "mandarin": lookup_mandarin(q),
            "timestamp": row["timestamp"],
            "timestamp_range": row["timestamp_range"],
            "start_ms": row["start_ms"],
            "tags": tags,
        })

    selected = select_evenly(candidates, TARGET_COUNT)
    if len(selected) < TARGET_COUNT:
        raise SystemExit(f"Only selected {len(selected)} rows; need {TARGET_COUNT}.")

    finalized = []
    for c in selected:
        tags = ensure_pragmatics(c["tags"], c["taigi"])
        cat = categorize_tags(tags)
        row = {
            **c,
            "tags": tags,
            **cat,
            "explanation": build_explanation(tags, c["taigi"]),
        }
        if not row["pinyin"]:
            row["pinyin"] = lookup_pinyin(row["taigi"]) or "（待補臺羅）"
        if not row["mandarin"]:
            row["mandarin"] = lookup_mandarin(row["taigi"]) or "（待補華語）"
        finalized.append(row)
    selected = finalized

    wb = Workbook()
    ws = wb.active
    ws.title = "EP1後半資料庫"

    headers = [
        "流水號", "時間戳記", "時間範圍", "台語", "臺羅", "臺灣華語",
        "句型結構", "詞彙與時態", "發音與變調", "語氣與情境",
        "文法Index(極簡名稱)", "文法Index_ID", "語言用法解析",
    ]
    header_fill = PatternFill("solid", fgColor="2D6A4F")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, c in enumerate(selected, 1):
        ws.append([
            i, c["timestamp"], c["timestamp_range"], c["taigi"], c["pinyin"],
            c["mandarin"], c["syntax"], c["semantics"], c["phonology"],
            c["pragmatics"], c["all_labels"], c["all_ids"], c["explanation"],
        ])

    widths = [6, 28, 18, 42, 48, 42, 22, 22, 22, 16, 36, 28, 80]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    wb.save(OUTPUT)

    missing_py = sum(1 for c in selected if c["pinyin"] == "（待補臺羅）")
    missing_zh = sum(1 for c in selected if c["mandarin"] == "（待補華語）")
    missing_prag = sum(1 for c in selected if not c["pragmatics"])
    times = [c["start_ms"] / 1000 for c in selected]

    print(f"Wrote {len(selected)} rows to {OUTPUT}")
    print(f"Time span: {min(times)//60:.0f}:{min(times)%60:02.0f} – {max(times)//60:.0f}:{max(times)%60:02.0f}")
    print(f"Missing 臺羅: {missing_py}, Missing 華語: {missing_zh}, Missing 語氣與情境: {missing_prag}")
    print("Time distribution (min):")
    for i, c in enumerate(selected):
        t = c["start_ms"] / 1000
        print(f"  {i+1:2d} {int(t//60):02d}:{int(t%60):02d} | {c['taigi'][:28]}")


if __name__ == "__main__":
    main()
