#!/usr/bin/env python3
"""Build EP1 Taiwanese subtitle database Excel from aligned SRT files."""

import json
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
TAIGI_SRT = BASE / "ep1前半段字幕台語字幕.srt"
MANDARIN_SRT = BASE / "ep1前半段字幕拜託.srt"
DRAMA_JS = BASE.parent / "drama-quotes-data.js"
OUTPUT = BASE / "台語資料庫_EP1.xlsx"
TARGET_COUNT = 50

LT_LABELS = {
    "kā-disposal": "「共」：處置",
    "kā-target": "「共」：對象",
    "hōo-passive": "「予」：被動",
    "hōo-causative": "「予」：使役",
    "comparative": "「較」：比較",
    "potential": "「袂」：可能",
    "extent-comp": "「甲」：程度",
    "conditionals": "「若」：假設",
    "aspect-perf": "「矣」：完成",
    "aspect-cont": "「咧」：持續",
    "aspect-neg": "「未」：否定",
    "modal-neg": "「莫」：禁止",
    "diminutive-a": "「仔」：指小",
    "directional": "「趨向」：複合",
    "neutral-tone": "「--」：輕聲",
    "contraction": "「合音」：緊縮",
    "reduplication": "「疊詞」：變調",
    "checked-tone": "「入聲」：急促音",
    "literary-coll": "「文白」：兩讀",
    "prag-angry": "嗆聲",
    "prag-sad": "哀怨",
    "prag-happy": "歡喜",
    "prag-sarcasm": "激刺",
    "prag-shock": "驚惶",
    "prag-idiom": "俗語",
}

SYNTAX_IDS = {
    "kā-disposal", "kā-target", "hōo-passive", "hōo-causative",
    "comparative", "potential", "extent-comp", "conditionals",
}
SEMANTICS_IDS = {
    "aspect-perf", "aspect-cont", "aspect-neg", "modal-neg",
    "diminutive-a", "directional",
}
PHONOLOGY_IDS = {
    "neutral-tone", "contraction", "reduplication", "checked-tone", "literary-coll",
}
PRAGMATICS_IDS = {
    "prag-angry", "prag-sad", "prag-happy", "prag-sarcasm", "prag-shock", "prag-idiom",
}

TAG_EXPLANATIONS = {
    "kā-disposal": "「共」為處置式標誌，將受事提前置於動詞前，表示對該對象施以動作。",
    "kā-target": "「共」引介動作對象，標示動作所指向的人或物。",
    "hōo-passive": "「予」表被動或遭受，主語承受他人施加的動作或結果。",
    "hōo-causative": "「予」表使役或給予，使對象做某事或得到某物。",
    "comparative": "「較／khah」表比較級，修飾形容詞或動詞，表示程度勝過他者。",
    "potential": "「袂」表否定可能，說明動作無法完成或辦不到。",
    "extent-comp": "「甲」為程度補語，連結動詞與結果狀態，表示做到某種程度。",
    "conditionals": "「若／nā」引導假設條件，後句說明條件成立時的結果。",
    "aspect-perf": "「矣」表完成貌，動作或狀態已經實現。",
    "aspect-cont": "「咧」表持續貌，動作正在進行或反覆發生。",
    "aspect-neg": "「未／猶未」表否定完成，事情尚未發生或尚未做完。",
    "modal-neg": "「莫／毋通」表禁止或勸阻，請對方不要做某事。",
    "diminutive-a": "「仔」為指小詞，使名詞帶有細小、親暱或口語色彩。",
    "directional": "趨向複合動詞（如轉去、出來）表動作的方向或位移。",
    "neutral-tone": "句末語助詞（如啊、喔、啦）常念輕聲，羅馬字以 -- 標示。",
    "contraction": "口語合音（如阮 guán、恁 lín、毋 m̄）為台語常見緊縮形式。",
    "reduplication": "疊字或疊詞（如緊咧緊咧）加強語氣或表持續、反覆。",
    "checked-tone": "入聲字（如急、法、賺、欲）發音短促，為台語八音特色之一。",
    "literary-coll": "同一漢字在口語與讀音層次有不同讀法，須依語境辨識。",
    "prag-angry": "語氣強硬、對立，常見於爭執、討債或衝突場景。",
    "prag-sad": "語氣哀怨、無奈，表達困境或失落。",
    "prag-happy": "語氣歡喜、熱絡，常見於招呼、慶祝或鼓勵。",
    "prag-sarcasm": "表面平淡實則諷刺，帶有反話或冷嘲意味。",
    "prag-shock": "表驚訝、疑惑或難以置信。",
    "prag-idiom": "諺語或慣用說法，承載文化常識與修辭傳統。",
}

# Curated 臺羅 for SRT lines (supplements drama-quotes-data.js matching)
PINYIN_OVERRIDES = {
    "人講出世著愛看八字時": "Lâng kóng chhut-sè tio̍h-ài khuànn pa̍t-jī--sî.",
    "不過有錢無錢著愛看自己": "Put-kò ū-tsînn bô-tsînn tio̍h-ài khuànn ka-tī.",
    "阮就是靠這个攤仔賺錢": "Guán tō-sī khò sì-kè thoaⁿ-á thàn-tsînn.",
    "照顧家庭，𤆬囡仔大漢": "Tsàu-kòo ka-tîng, tshuā gín-á tuā-hàn.",
    "阮就是靠夜市這葩燈。": "Guán tō-sī khò iā-tshī tsit pha tîng.",
    "大家來予我湊熱鬧": "Ta-ke lâi hōo guá tshàu jia̍t-lāu.",
    "緊咧、緊咧，轉去作生理": "Kín--leh, kín--leh, tńg-khì tsò-sing-lí.",
    "趁大錢哦，緊咧！": "Thàn tuā-tsînn--ooh, kín--leh!",
    "若來我小野貓的Pub捧場": "Nā lâi guá Sió-iá-bâu ê Pub phòng-tiûnn.",
    "予我一杯血腥瑪麗": "Hōo guá tsi̍t poe hiat-hing-má-lī.",
    "阿珠、阿花，緊咧！": "A-tsu, A-hue, kín--leh!",
    "就真正有較困難啦": "Tō tsin-tsiànn ū khah khùn-lân--lah.",
    "啊無恁是想欲按怎？": "Ah bô lín sī siūnn-beh án-nuá?",
    "你毋顧阮的死活。": "Lí m̄-kòo guán ê sí-ua̍h.",
    "阮嘛欲討你的命。": "Guán mā beh thó lí ê miā.",
    "莫走！": "Mài tsáu!",
    "按怎矣？": "Ánn-nuá--ah?",
    "發生啥物代誌？": "Hoat-seng siáⁿ-mih tāi-tsì?",
    "恁咧創啥？": "Lín leh tshòng siáⁿ?",
    "為啥物欲拍伊？": "Ūi-siáⁿ beh phah i?",
    "伊毋發薪水予阮。": "I m̄ hoat suí-sìn hōo guán.",
    "故意欲予阮餓死。": "Kòo-ì beh hōo guán gō-sí.",
    "我就共你講過矣": "Guá tō kā lí kóng kòe--ah.",
    "大陸的錢若匯過來，我隨發予恁。": "Tāi-lio̍k ê tsînn nā huē kuè-lâi, guá suî hoat hōo lín.",
    "毋通閣烏白講矣。": "M̄-thang koh oo-pe̍h kóng--ah.",
    "過去你共阮的年終獎金……": "Kòe-khì lí kā guán ê nî-tiong tsióng-kim……",
    "佮一寡存款，攏騙去大陸投資。": "Kah tsi̍t-kó tsûn-khuǹ, lóng phian khì Tāi-lio̍k tâu-tsu.",
    "今年攏欠四、五個月的薪水。": "Kin-nî lóng khiam sì, gōo kò goe̍h ê suí-sìn.",
    "阮等你大陸匯錢過來。": "Guán tán lí Tāi-lio̍k huē-tsînn kuè-lâi.",
    "阮歸氣等反攻大陸": "Guán kui-khì tán hoán-kong Tāi-lio̍k.",
    "才閣共你領，好無？": "Tsiah koh kā lí líng, hó--bô?",
    "若李慶祥毋發薪水": "Nā Lí Khìng-siông m̄ hoat suí-sìn.",
    "阮欲按怎？": "Guán beh án-nuá?",
    "漢良，慶祥到底欠恁偌濟錢？": "Hàn-liông, Khìng-siông tàu-té khiam lín gōa-tsuē tsînn?",
    "閣有我借伊的。": "Koh ū guá chai i ê.",
    "總共一百二十萬。": "Tsóng-kiong tsi̍t-pah-jī-tsa̍p bān.",
    "欲按怎還？": "Beh án-nuá huan?",
    "免還啦。": "Bián huan--lah.",
    "欲跤欲手，隨在𪜶。": "Beh kha beh tshiú, suî tī in.",
    "斷手斷跤，事情就會當解決無？": "Tn̄g-tshiú tn̄g-kha, sū-tsì tō ē-tàng kái-kuat bô?",
    "解決袂了，也無法度啊。": "Kái-kuat bē-liáu, iā bô-hoat-tōo--ah.",
    "歸氣按呢。": "Kui-khì án-ne.",
    "我的頭剁落來予恁當椅仔坐啦": "Guá ê thâu to̍h-lo̍h-lâi hōo lín tòng í-á tsē--lah.",
    "本來就是按呢啊。": "Pún-lâi tō-sī án-ne--ah.",
    "雖然講我佇外口有家庭，毋過遮爾多年來，我嘛是較照顧你兩個囡仔啊。": "Suî-jiân kóng guá tī guā-kháu ū ka-tîng, m̄-koh tsiah-nī tsuē nî lâi, guá mā-sī khah tsàu-kòo lí nn̄g ê gín-á--ah.",
    "我佇大陸這款拚生拋死，也是欲拚予恁過較好的日子。": "Guá tī Tāi-lio̍k tsit-khuán piànn-sinn-piànn-sí, iā-sī beh piànn hōo lín kuè khah hó ê ji̍t-tsí.",
    "我就運途䆀啊，做成衣賺的錢喔，去賭股票、炒期貨。": "Guá tō sī ūn-tôo bái--ah, tsò siânn-i thàn ê tsînn--ooh, khì tóo kóo-phiò, tshá kî-huò.",
    "做一回喔，這馬賺的錢都全了了。": "Tsò tsi̍t-huê--ooh, tsit-má thàn ê tsînn lóng tsuân liáu-liáu.",
    "順其自然就好。": "Sūn-kî tsì-jiân tō hó.",
    "你若是無法度，阮就共你處理。": "Lí nā-sī bô-huat-tōo, guán tō kā lí tshú-lí.",
    "你敢有聽我講？": "Lí kám-ū thiann guá kóng?",
    "我講的攏是事實啦。": "Guá kóng ê lóng sī sū-si̍t--lah.",
    "你毋是講大陸的錢會匯來？": "Lí m̄-sī kóng Tāi-lio̍k ê tsînn ē huē-lâi?",
    "到今猶未匯一塊錢來。": "Kàu-tann iáu-buē huē tsi̍t-tè tsînn lâi.",
    "你閣欲騙阮幾擺？": "Lí koh beh phian guán kuí-pái?",
    "阮已經袂相信你了。": "Guán í-king bē siong-sìn lí--ah.",
    "你若閣講，我就叫警察來。": "Lí nā koh kóng, guá tō kiò kíng-tshat lâi.",
    "你敢拍伊？": "Lí kám phah i?",
    "伊欠阮的錢，阮欲討回來。": "I khiam guán ê tsînn, guán beh thó--huê-lâi.",
    "你莫摃伊啦！": "Lí mài lòng i--lah!",
    "按呢會出人命啦！": "Án-ne ē tshut-lâng-miā--lah!",
    "大家冷靜一擺！": "Ta-ke líng-tsīng tsi̍t-pái!",
    "有話好講，毋通拍人。": "Ū-uē hó kóng, m̄-thang phah-lâng.",
    "你講的毋著。": "Lí kóng ê m̄-tio̍h.",
    "阮攏知影你咧騙人。": "Guán lóng tsai-iánn lí leh phian-lâng.",
    "錢若無還，阮就告你。": "Tsînn nā bô huan, guán tō kòo lí.",
    "你敢會當一個月內還清？": "Lí kám ē-tàng tsi̍t kò goe̍h lāi huan-tshing?",
    "我會盡力，毋過真困難。": "Guá ē tsìn-li̍k, m̄-koh tsin khùn-lân.",
    "你予阮等幾若個月矣！": "Lí hōo guán tán kuí-ji̍t kò goe̍h--ah!",
    "阮袂閣等矣！": "Guán bē koh tán--ah!",
    "你若袂還，阮就共你提告。": "Lí nā bē huan, guán tō kā lí the̍h-kòo.",
    "你敢知影阮幾艱苦？": "Lí kám tsai-iánn guán kuí kan-khóo?",
    "囡仔的學費都未繳。": "Gín-á ê ha̍k-huì lóng buē kiàu.",
    "厝的貸款猶未還。": "Tshù ê thài-khuǎn iáu-buē huan.",
    "你講會匯，結果攏無。": "Lí kóng ē huē, kiat-kó lóng bô.",
    "阮已經忍幾若年矣。": "Guán í-king jím kuí-ji̍t nî--ah.",
    "你敢欲予阮餓死？": "Lí kám beh hōo guán gō-sí?",
    "毋通再講假話！": "M̄-thang tsài kóng ké-uē!",
    "你欠阮的，阮欲討回來！": "Lí khiam guán ê, guán beh thó--huê-lâi!",
    "你敢有誠心欲還無？": "Lí kám-ū sîng-sim beh huan bô?",
    "你若誠心，阮就閣等你一擺。": "Lí nā sîng-sim, guán tō koh tán lí tsi̍t-pái.",
    "你若是閣拖，阮就袂客氣。": "Lí nā-sī koh thoa, guán tō bē kheh-khì.",
    "阮已經袂閣相信你了。": "Guán í-king bē koh siong-sìn lí--ah.",
    "你敢知影阮為啥欲拍你？": "Lí kám tsai-iánn guán ūi-siáⁿ beh phah lí?",
    "因為你騙阮的錢！": "In-uī lí phian guán ê tsînn!",
    "你講大陸有錢，結果攏是空空的。": "Lí kóng Tāi-lio̍k ū tsînn, kiat-kó lóng sī khang-khang ê.",
    "阮等你等到頭毛白矣！": "Guán tán lí tán kàu thâu-mn̂g pe̍h--ah!",
    "你敢有想欲按怎解決？": "Lí kám-ū siūnn-beh án-nuá kái-kuat?",
    "你若袂還，阮就共你見官。": "Lí nā bē huan, guán tō kā lí kiàn-kuan.",
    "你敢會當這禮拜內先還一半？": "Lí kám ē-tàng tsit lé-pài lāi seng huan tsi̍t-puànn?",
    "我會想辦法，毋過需要時間。": "Guá ē siūnn-pān-huat, m̄-koh su-iàu sî-kan.",
    "你閣拖，阮就袂閣等。": "Lí koh thoa, guán tō bē koh tán.",
    "你敢欲予阮全家餓死？": "Lí kám beh hōo guán tsuan-ka gō-sí?",
    "阮已經忍無可忍矣！": "Guán í-king jím bô-khó-jím--ah!",
    "你若閣騙，阮就報警。": "Lí nā koh phian, guán tō pò-kíng.",
    "你敢有聽阮講？": "Lí kám-ū thiann guán kóng?",
    "阮講的攏是事實！": "Guán kóng ê lóng sī sū-si̍t!",
    "你欠阮一百二十萬！": "Lí khiam guán tsi̍t-pah-jī-tsa̍p bān!",
    "你敢欲賴帳無？": "Lí kám beh lāi-tiūnn bô?",
    "阮袂予你賴帳！": "Guán bē hōo lí lāi-tiūnn!",
    "你若誠意，阮就閣商量。": "Lí nā sîng-ì, guán tō koh siong-lióng.",
    "你若是閣拖，阮就提告。": "Lí nā-sī koh thoa, guán tō the̍h-kòo.",
    "你敢知影阮幾絕望？": "Lí kám tsai-iánn guán kuí tsua̍t-bōng?",
    "囡仔欲交學費，阮無錢。": "Gín-á beh kau ha̍k-huì, guán bô tsînn.",
    "你講會幫阮，結果攏無。": "Lí kóng ē pang guán, kiat-kó lóng bô.",
    "阮已經袂閣信你了。": "Guán í-king bē koh sìn lí--ah.",
    "你敢欲予阮走投無路？": "Lí kám beh hōo guán tsáu-thâu bô-lōo?",
    "毋通再講空話！": "M̄-thang tsài kóng khang-uē!",
    "你欠的錢，阮欲討回來！": "Lí khiam ê tsînn, guán beh thó--huê-lâi!",
    "你敢有誠心無？": "Lí kám-ū sîng-sim bô?",
    "你若誠心，阮就閣等。": "Lí nā sîng-sim, guán tō koh tán.",
    "你若是閣騙，阮就袂客氣。": "Lí nā-sī koh phian, guán tō bē kheh-khì.",
    "阮已經忍幾若年矣！": "Guán í-king jím kuí-ji̍t nî--ah!",
    "你敢知影阮為啥欲討？": "Lí kám tsai-iánn guán ūi-siáⁿ beh thó?",
    "因為你騙阮！": "In-uī lí phian guán!",
    "你講有錢，結果攏是假。": "Lí kóng ū tsînn, kiat-kó lóng sī ké.",
    "阮等你等到心冷矣！": "Guán tán lí tán kàu sim líng--ah!",
    "你敢有想欲按怎？": "Lí kám-ū siūnn-beh án-nuá?",
    "你若袂還，阮就見官。": "Lí nā bē huan, guán tō kiàn-kuan.",
    "你敢會當先還一半？": "Lí kám ē-tàng seng huan tsi̍t-puànn?",
    "我會盡力，毋過真困難。": "Guá ē tsìn-li̍k, m̄-koh tsin khùn-lân.",
    "你閣拖，阮就袂等。": "Lí koh thoa, guán tō bē tán.",
    "你敢欲予阮餓死？": "Lí kám beh hōo guán gō-sí?",
    "阮已經忍無可忍！": "Guán í-king jím bô-khó-jím!",
    "你若閣騙，阮就報警。": "Lí nā koh phian, guán tō pò-kíng.",
    "這樣事情比較快解決啦": "Án-ne sū-tsì pí-kàu kín kái-kuat--lah.",
    "啊無按呢，予我三工。": "Ah bô án-ne, hōo guá sann-kang.",
    "三工拄好予恁收行李走人。": "Sann-kang tú-hó hōo lín siu hêng-lí tsáu-lâng.",
    "若是寄望這一攤賺來還": "Nā-sī kià-bāng tsit tê thàn lâi huan.",
    "阮乾脆拍攤較甘心": "Guán tshian-tshui phah-thoaⁿ khah kam-sim.",
    "毋通掠攤仔。": "M̄-thang lia̍h-thoaⁿ-á.",
    "想欲拍歹攤仔，恁先顧家己。": "Siūnn-beh phah pháiⁿ-thoaⁿ-á, lín seng kòo ka-kī.",
    "毋准任何人惹代誌。": "M̄-tsún jīm-hô lâng jiá tāi-tsì.",
    "欠儂錢緊共儂處理": "Khiàm lâng-tsînn kín kā lâng tshú-lí.",
    "我來揣伊，袂使嗎？": "Guá lâi tshuē i, bē-sái--bô?",
    "共儂處理啊啦！": "Kā lâng tshú-lí--ah--lah!",
    "好矣，無代誌矣。": "Hó--ah, bô tāi-tsì--ah.",
    "毋通圍佇遮看鬧熱。": "M̄-thang ûi tī tsia khuànn lāu-jia̍t.",
    "賺錢較重要啦": "Thàn-tsînn khah tiōng-iàu--lah.",
    "這張是我頂學期得著的獎狀喔。": "Tsit tiuⁿ sī guá tíng ha̍k-kî tit-tio̍h ê tsióng-tsuāng--ooh.",
    "若準大陸的生理做袂落去……": "Nā tsún Tāi-lio̍k ê sing-lí tsò bē-lo̍h-khì……",
    "講過好幾擺矣，攏講袂聽。": "Kóng kòe hó-kuí pái--ah, lóng kóng bē thiann.",
    "囡仔欲按怎讀冊？": "Gín-á beh án-nuá tha̍k-tsheh?",
    "講有代誌欲參詳啦。": "Kóng ū tāi-tsì beh tsham-siông--lah.",
    "莫閣悶悶無出聲啦！": "Mài koh būn-būn bô tshut-siann--lah!",
    "我就欲共伊冰起來。": "Guá tō beh kā i ping--khí-lâi.",
    "請問這位太太，妳有啥物貴事？": "Tshiánn-mn̄g tsit uī thài-thài, lí ū siáⁿ-mih kuì-sū?",
    "逐家攏轉去做生理。": "Ta̍k-ke lóng tńg-khì tsò sing-lí.",
    "你佮你翁的代誌，": "Lí kah lí ang ê tāi-tsì,",
    "我們是在夜市做生意": "Guán tī iā-tshī tsò-sing-lī.",
    "感恩啊": "Kám-un--ah.",
    "唱歌免費喔": "Tshàng-koa bián-huì--ooh.",
    "來……到旁邊說": "Lâi……kàu pîng-pinn kóng.",
    "救命啊": "Kiù-miā--ah.",
    "不管我們死活": "Put-kóan guán ê sí-ua̍h.",
    "恁攏聽無懂嗎？": "Lín lóng thiann bô tháng--bô?",
    "斷跤斷手": "Tn̄g-kha tn̄g-tshiú.",
    "遐爾濟啊！": "Hiah-nī tsē--ah!",
    "我還有這攤藥膳排骨要顧啦": "Guá iáu ū tsit thoaⁿ io̍h-siàn pâi-kut beh kòo--lah.",
    "愈弄愈麻煩。": "Jū lòng jū mâ-hoan.",
    "尻脊骿癢甲欲死。": "Kha-tí-kuann jiánn kah beh sí.",
    "總有手頭較絚的時候。": "Tsóng-ū tshiú-thâu khah ké ê sî-hāu.",
    "生理若䆀，我看歸氣共公司頂讓予別人做。": "Sing-lí nā bái, guá khuànn kui-khì kā kong-si téng-jiōng hōo pa̍t-lâng tsò.",
    "也是無法度解決。": "Iā-sī bô-hoat-tōo kái-kuat.",
    "好食的玉米喔。": "Hó-tsia̍h ê gôk-mî--ooh.",
    "乾淨溜溜": "Kan-tsing liu-liu.",
    "若無夠會使加湯。": "Nā bô kàu ē-sái ke thng.",
    "恁會當共伊討錢": "Lín ē-tàng kā i thó tsînn.",
    "轉去好好的參詳。": "Tńg-khì hó-hó ê tsham-siông.",
    "我欲予阿爸看我的獎狀。": "Guá beh hōo a-pa khuànn guá ê tsióng-tsuāng.",
    "我佮你睏，好無？": "Guá kah lí khùn, hó--bô?",
    "予伊爛啦。": "Hōo i nūn--lah.",
    "無欲按怎？": "Bô beh án-nuá?",
    "我若無食，伊就會疑神疑鬼。": "Guá nā bô tsia̍h, i tō ē gî-sîn gî-kuí.",
    "會當去洗身軀矣。": "Ē-tàng khì sé sin-khu--ah.",
    "生理做甲按呢，": "Sing-lí tsò kah án-ne,",
    "你按怎講，咱就按怎聽。": "Lí án-nuá kóng, lán tō án-nuá thiann.",
    "按呢就對矣嘛。": "Án-ne tō tùi--ah--má.",
    "多謝逐家的鬥相共。": "Tō-siā ta̍k-ke ê tàu-saⁿ-kāng.",
    "歸氣攏共伊賭落去。": "Kui-khì lóng kā i tó lo̍h-khì.",
    "拄仔好。": "Tú-á-hó.",
    "阮攏用當歸、": "Guán lóng iōng tong-kui、",
    "桌仔椅仔遮爾拉撒，": "Toh-á í-á tsiah-nī lann-sat,",
    "欲按怎坐？": "Beh án-nuá tsē?",
    "看伊氣質真好，": "Khuànn i khì-tsit tsin hó,",
    "來，捧場一下。": "Lâi, phòng-tiûnn tsi̍t-ē.",
}

CHECKED_TONE_PAT = re.compile(
    r"急|法|賺|結|發|質|職|值|識|歇|入聲|劫|壓|習|逐|逐|逐|逐"
)


def normalize_text(s: str) -> str:
    s = s.strip()
    s = re.sub(r"[…\.。，,、！!？?「」『』""\"'\s]+", "", s)
    return s


def parse_srt(path: Path) -> list[dict]:
    """Parse SRT even when blank lines between cues are missing."""
    lines = path.read_text(encoding="utf-8").splitlines()
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
        time_line = lines[i].strip()
        if "-->" not in time_line:
            i += 1
            continue
        i += 1
        text_lines = []
        while i < len(lines):
            s = lines[i].strip()
            if not s:
                i += 1
                break
            if re.match(r"^\d+$", s) and i + 1 < len(lines) and "-->" in lines[i + 1]:
                break
            text_lines.append(s)
            i += 1
        m = re.match(
            r"(\d{2}):(\d{2}):(\d{2}),(\d{3})\s*-->\s*(\d{2}):(\d{2}):(\d{2}),(\d{3})",
            time_line,
        )
        if not m or not text_lines:
            continue
        sh, sm, ss, sms, eh, em, es, ems = map(int, m.groups())
        start_ms = ((sh * 60 + sm) * 60 + ss) * 1000 + sms
        end_ms = ((eh * 60 + em) * 60 + es) * 1000 + ems
        text = " ".join(text_lines).strip()
        if not text or text.upper() in {"OK", "OK."}:
            continue
        entries.append({
            "start_ms": start_ms,
            "end_ms": end_ms,
            "timestamp": time_line.replace(" ", ""),
            "timestamp_range": f"{sh:02d}:{sm:02d}:{ss:02d}-{eh:02d}:{em:02d}:{es:02d}",
            "text": text,
        })
    return entries


def load_drama_pinyin() -> dict[str, str]:
    if not DRAMA_JS.exists():
        return {}
    content = DRAMA_JS.read_text(encoding="utf-8")
    mapping = {}
    quotes = re.findall(r'quote:\s*"([^"]+)"', content)
    pinyins = re.findall(r'pinyin:\s*"([^"]+)"', content)
    for q, p in zip(quotes, pinyins):
        mapping[normalize_text(q)] = p
        mapping[q.strip()] = p
    return mapping


def infer_pragmatics(q: str) -> set[str]:
    tags = set()
    if re.search(r"出世|八字|人講|俗諺|諺語|點亮人生", q):
        tags.add("prag-idiom")
    if re.search(
        r"感謝|感恩|湊熱鬧|趁大錢|緊咧|生意|捧場|免費|招呼|歡迎|拍拚|夕陽|日出",
        q,
    ):
        tags.add("prag-happy")
    if re.search(
        r"拍|討命|討債|討錢|毋顧|莫走|莫|摃|提告|見官|報警|袂客氣|烏白講|騙|欠|"
        r"斷手|斷跤|法院|告你|追拍|白賊|欺騙|憑啥|超過|搶人|拖出去|狐狸精",
        q,
    ):
        tags.add("prag-angry")
    if re.search(
        r"困難|運途|哀怨|無法度|餓死|絕望|心冷|艱苦|死活|債務|周轉|開了了|運氣歹|"
        r"煩惱|無睏|忍一下",
        q,
    ):
        tags.add("prag-sad")
    if re.search(r"按呢|順其自然|歸氣按呢|本來就是|毋是按呢", q):
        tags.add("prag-sarcasm")
    if re.search(r"啥物|按怎|蛤|敢有|創啥|代誌|好無|敢會|不知影", q):
        tags.add("prag-shock")
    return tags


def ensure_pragmatics(tags: list[str], q: str) -> list[str]:
    """Guarantee at least one pragmatics index for searchability."""
    tag_set = set(tags)
    if tag_set & PRAGMATICS_IDS:
        return sorted(tag_set)

    tag_set |= infer_pragmatics(q)
    if tag_set & PRAGMATICS_IDS:
        return sorted(tag_set)

    if re.search(r"[？?]|按怎|好無|敢有|啥物|蛤", q):
        tag_set.add("prag-shock")
    elif re.search(r"阮|恁|討|拍|莫|毋|騙|欠|法院", q):
        tag_set.add("prag-angry")
    elif re.search(r"困難|死活|餓|艱苦|無法度|債", q):
        tag_set.add("prag-sad")
    elif re.search(r"感謝|賺|生意|熱鬧|緊咧|捧場", q):
        tag_set.add("prag-happy")
    elif re.search(r"按呢|順其|本來", q):
        tag_set.add("prag-sarcasm")
    elif re.search(r"人講|出世|八字", q):
        tag_set.add("prag-idiom")
    else:
        tag_set.add("prag-happy")
    return sorted(tag_set)


def infer_linguistic_tags(q: str, pinyin: str = "") -> list[str]:
    tags: set[str] = set()
    g = ""
    p = pinyin or ""

    if "共" in q and re.search(r"拖|拍|弄|提|搶|聽|處理|講|領|扛|見官|提告", q):
        tags.add("kā-disposal")
    if re.search(r"共[你伊我佮]|共阮|共恁|共儂|共你|共阮", q):
        tags.add("kā-target")
    if "予" in q:
        if re.search(r"予[我你伊佮阮恁儂]|當椅仔|湊熱鬧|發予|剁落來予", q):
            tags.add("hōo-causative")
        if re.search(r"被|叫|拍|講|領|賭|欠|欺負|餓死", q):
            tags.add("hōo-passive")
    if re.search(r"較|愈.*愈|khah", q + g + p, re.I):
        tags.add("comparative")
    if re.search(r"袂|會得|ē-tit|buē|袂了|袂閣|袂等|袂相信|袂予", q + g + p, re.I):
        tags.add("potential")
    if "甲" in q or re.search(r"程度|結果補語", g):
        tags.add("extent-comp")
    if re.search(r"若|假使|欲.*著愛|nā|ká-sú|beh|若是", q + g + p, re.I):
        tags.add("conditionals")
    if re.search(r"矣|清矣|過去矣|講過矣|白矣|冷矣|忍.*矣|無代誌矣|好幾擺矣", q):
        tags.add("aspect-perf")
    if re.search(r"咧|--leh|teh|緊咧", q + g + p, re.I):
        tags.add("aspect-cont")
    if re.search(r"未|猶未|buē|猶未匯|都未", q + g + p, re.I):
        tags.add("aspect-neg")
    if re.search(r"莫|袂使|mài|buē-sái|毋通", q + g + p, re.I):
        tags.add("modal-neg")
    if re.search(r"仔|gín-á|-á|攤仔|椅仔|囡仔", q + g + p, re.I):
        tags.add("diminutive-a")
    if re.search(r"出來|轉來|轉去|tshut-lâi|tńg-lâi|tńg-khì|趨向|討回來|匯過來|剁落來", q + g + p, re.I):
        tags.add("directional")
    if re.search(r"[啊喔啦哦耶].{0,2}$|--", q + p) or re.search(r"輕聲", g):
        tags.add("neutral-tone")
    if re.search(r"合音|buāi|gún|lín|阮|恁|毋|佮|m̄", q + g + p, re.I):
        tags.add("contraction")
    if re.search(r"(.)\1|疊|liáu-liáu|緊咧、緊咧", q) or re.search(r"疊詞|疊字", g):
        tags.add("reduplication")
    if re.search(
        r"入聲|tsia̍t|kip|ki̍p|kiat|huat|sip|si̍t|jia̍t|liáu|bē|beh|tio̍h",
        p + g,
        re.I,
    ) or CHECKED_TONE_PAT.search(q):
        tags.add("checked-tone")
    if re.search(r"文白|兩讀|ji̍t|著愛|欲|tio̍h-ài", g + p + q, re.I):
        tags.add("literary-coll")

    tags |= infer_pragmatics(q)
    return sorted(tags)


def lookup_pinyin(q: str, drama_map: dict[str, str]) -> str:
    if q in PINYIN_OVERRIDES:
        return PINYIN_OVERRIDES[q]
    if q in drama_map:
        return drama_map[q]
    nq = normalize_text(q)
    if nq in drama_map:
        return drama_map[nq]
    for dk, dp in drama_map.items():
        if nq in dk or dk in nq:
            if len(dk) >= 4:
                return dp
    return ""


def build_explanation(tags: list[str], q: str) -> str:
    parts = []
    for tid in tags:
        base = TAG_EXPLANATIONS.get(tid, "")
        if base:
            parts.append(f"【{LT_LABELS[tid]}】{base}")
    if not parts:
        return ""
    return " ".join(parts)


def align_subtitles(taigi: list[dict], mandarin: list[dict]) -> list[dict]:
    mand_by_start = {e["start_ms"]: e for e in mandarin}
    mand_sorted = sorted(mandarin, key=lambda x: x["start_ms"])
    rows = []
    for t in taigi:
        mand = mand_by_start.get(t["start_ms"])
        if not mand:
            for m in mand_sorted:
                if abs(m["start_ms"] - t["start_ms"]) <= 500:
                    mand = m
                    break
        rows.append({
            **t,
            "mandarin": mand["text"] if mand else "",
        })
    return rows


def categorize_tags(tags: list[str]) -> dict[str, str]:
    def labels_for(id_set):
        return "、".join(LT_LABELS[t] for t in tags if t in id_set)

    prag = labels_for(PRAGMATICS_IDS)
    return {
        "syntax": labels_for(SYNTAX_IDS),
        "semantics": labels_for(SEMANTICS_IDS),
        "phonology": labels_for(PHONOLOGY_IDS),
        "pragmatics": prag,
        "all_labels": "、".join(LT_LABELS[t] for t in tags),
        "all_ids": "、".join(tags),
    }


def score_candidate(c: dict) -> tuple:
    return (len(c["tags"]), len(c["taigi"]), -c["start_ms"])


def select_evenly(candidates: list[dict], target: int, t0: float, t1: float) -> list[dict]:
    """Pick one sentence per time bin for even coverage."""
    bin_w = (t1 - t0) / target
    used = set()
    selected = []

    for i in range(target):
        lo = t0 + i * bin_w
        hi = t0 + (i + 1) * bin_w
        center = (lo + hi) / 2
        pool = [c for c in candidates if normalize_text(c["taigi"]) not in used]
        in_bin = [c for c in pool if lo <= c["start_ms"] / 1000 < hi]
        if not in_bin:
            in_bin = sorted(pool, key=lambda c: abs(c["start_ms"] / 1000 - center))[:8]
        if not in_bin:
            continue
        pick = max(in_bin, key=score_candidate)
        selected.append(pick)
        used.add(normalize_text(pick["taigi"]))

    return sorted(selected, key=lambda c: c["start_ms"])


def finalize_candidate(c: dict, drama_map: dict) -> dict:
    tags = ensure_pragmatics(c["tags"], c["taigi"])
    cat = categorize_tags(tags)
    c = {**c, "tags": tags, **cat, "explanation": build_explanation(tags, c["taigi"])}
    c["pinyin"] = lookup_pinyin(c["taigi"], drama_map) or "（待補臺羅）"
    return c


def main():
    taigi = parse_srt(TAIGI_SRT)
    mandarin = parse_srt(MANDARIN_SRT)
    drama_map = load_drama_pinyin()
    aligned = align_subtitles(taigi, mandarin)

    candidates = []
    for row in aligned:
        q = row["text"]
        if len(q) < 3:
            continue
        if re.fullmatch(r"[走救]+", q):
            continue
        if q in {"好", "是啊", "來", "人客", "招呼一下", "不要跑"}:
            continue

        pinyin = lookup_pinyin(q, drama_map)
        tags = infer_linguistic_tags(q, pinyin)
        if not tags:
            continue
        if pinyin:
            tags = sorted(set(tags) | set(infer_linguistic_tags(q, pinyin)))

        candidates.append({
            "taigi": q,
            "pinyin": pinyin,
            "mandarin": row["mandarin"],
            "timestamp": row["timestamp"],
            "timestamp_range": row["timestamp_range"],
            "start_ms": row["start_ms"],
            "tags": tags,
        })

    if not candidates:
        raise SystemExit("No tagged candidates found.")

    t0 = min(c["start_ms"] for c in candidates) / 1000
    t1 = max(c["start_ms"] for c in taigi) / 1000
    selected = select_evenly(candidates, TARGET_COUNT, t0, t1)

    if len(selected) < TARGET_COUNT:
        raise SystemExit(f"Only selected {len(selected)} rows; need {TARGET_COUNT}.")

    selected = [finalize_candidate(c, drama_map) for c in selected]

    wb = Workbook()
    ws = wb.active
    ws.title = "EP1台語資料庫"

    headers = [
        "流水號",
        "時間戳記",
        "時間範圍",
        "台語",
        "臺羅",
        "臺灣華語",
        "句型結構",
        "詞彙與時態",
        "發音與變調",
        "語氣與情境",
        "文法Index(極簡名稱)",
        "文法Index_ID",
        "語言用法解析",
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
            i,
            c["timestamp"],
            c["timestamp_range"],
            c["taigi"],
            c["pinyin"],
            c["mandarin"],
            c["syntax"],
            c["semantics"],
            c["phonology"],
            c["pragmatics"],
            c["all_labels"],
            c["all_ids"],
            c["explanation"],
        ])

    widths = [6, 28, 18, 42, 48, 42, 22, 22, 22, 16, 36, 28, 80]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"
    wb.save(OUTPUT)

    missing_pinyin = sum(1 for c in selected if c["pinyin"] == "（待補臺羅）")
    missing_prag = sum(1 for c in selected if not c["pragmatics"])
    times = [c["start_ms"] / 1000 for c in selected]
    print(f"Wrote {len(selected)} rows to {OUTPUT}")
    print(f"Time span: {int(min(times)//60)}:{int(min(times)%60):02d} – {int(max(times)//60)}:{int(max(times)%60):02d}")
    print(f"Missing 臺羅: {missing_pinyin}, Missing 語氣與情境: {missing_prag}")
    for i, c in enumerate(selected, 1):
        t = c["start_ms"] / 1000
        print(f"  {i:2d} {int(t//60):02d}:{int(t%60):02d} [{c['pragmatics']}] {c['taigi'][:24]}")
    tag_counts = {}
    for c in selected:
        for t in c["tags"]:
            tag_counts[t] = tag_counts.get(t, 0) + 1
    print("Tag distribution:", json.dumps(tag_counts, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
