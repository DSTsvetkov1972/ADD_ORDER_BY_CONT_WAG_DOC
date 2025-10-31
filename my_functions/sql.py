def sql (doc_cont_vag):
    return f"""
WITH
YOUR_EXCEL AS (
SELECT
--МЕЖДУ ОДИНАРНЫХ КОВЫЧЕК ПОДСТАВЬТЕ `№ платформы`,`№ отправки`,`№ контейнера` ИЗ ЭКСЕЛЬ
/*'
37785506|TKRU4286237|94207792
37785502|TKRU4609474|94207792
' AS `Сцеп`*/
'
{ doc_cont_vag }
' AS `Сцеп`
--) SELECT * FROM YOUR_EXCEL
),
YOUR_EXCEL AS (
SELECT
	arrayJoin(splitByChar('\r', replace(`Сцеп`, '\n', ''))) AS `Номер накладной|№ Контейнера|№ Вагона`
FROM
	YOUR_EXCEL
WHERE
	length(`Номер накладной|№ Контейнера|№ Вагона`)>2
--) SELECT * FROM YOUR_EXCEL
),
YOUR_EXCEL AS (
SELECT
	`Номер накладной|№ Контейнера|№ Вагона`,
	(splitByChar('|', `Номер накладной|№ Контейнера|№ Вагона`) AS parts)[1] AS `Номер накладной`,
	parts[2] AS `№ Контейнера`,
	parts[3] AS `№ Вагона`
FROM
	YOUR_EXCEL
--) SELECT * FROM YOUR_EXCEL WHERE `№ Вагона`='94207792'
),
ETRAN_CAR AS (
SELECT
	--this_car_id,
	invoiceid,--carclaimid,carclaimnumotp,carclaimnumpod,cartypeid,cartypecode,cartypename,
	carnumber--,
	--carorder--,carownercountrycode,carownercountryname,carownertypeid,carownertypename,carownerid,carownerokpo,carownername,cartenantid,cartenantokpo,cartenantname,carpocketcount,carplacescount,cartonnage,caraxles,carvolume,carweightdep,carweightdepreal,carweightgross,carweightnet,carweightadddev,caradddevwithgoods,carpriorfreightcode,carpriorfreightname,carguidecount,caroutsizecode,carframeweight,carframewagnum,cartopheight,carmainshtabquantity,carmainshtabheight,carheadshtabquantity,carforestvolume,carliquidtemperature,carliquidheight,carliquiddensity,carliquidvolume,cartanktype,carrefnum,carrefcount,carrolls,carconnectcode,cariscover,carsign,carlength,cardepnormdocid,cardepnormdocname,cardeppart,cardeparc,cardepsec,cardepcond,carntu_mtu_id,carntu_mtunumber,carntu_mtudate,carntuclearance,carmtuscheme
FROM 
	itrans__etran_invoice_car AS ETRAN_CAR
	INNER JOIN (SELECT DISTINCT `№ Вагона` FROM YOUR_EXCEL) AS YOUR_EXCEL ON YOUR_EXCEL.`№ Вагона`=ETRAN_CAR.carnumber
--) SELECT * FROM ETRAN_CAR --WHERE carnumber = '91778704'
),
ETRAN_CONT AS (
SELECT DISTINCT
	--this_cont_id,
	invoiceid,--contclaimnumotp,contclaimnumpod,
	contnumber--,--conttonnageid,conttonnage,
	--contcarorder--,conttypebig,conttypebigname,contsizebig,contwidthfoot,contweightdep,contpocketcount,contplacescount,contweightgross,contweightnet,contvolume,contownercountrycode,contownercountryname,contownertypeid,contownertypename,contownerid,contownerokpo,contownername
FROM 
	itrans__etran_invoice_cont AS ETRAN_CONT
	INNER JOIN (SELECT DISTINCT `№ Контейнера` FROM YOUR_EXCEL) AS YOUR_EXCEL ON YOUR_EXCEL.`№ Контейнера`=ETRAN_CONT.contnumber
--) SELECT * FROM ETRAN_CONT --WHERE invoiceid='1661376141' ORDER BY  invoiceid DESC
),
ETRAN AS ( -- SELECT * FROM etran_invoice
SELECT
	invoiceid,--invunp,
	invdatecreate,
	max(invdatecreate) OVER (PARTITION BY invnumber, invdatedeparture) AS max_invdatecreate,	
	--invdatepres,invoicestateid,--
	invoicestate,--invlastoper,invneedforecp,invecpsign,invtypeid,invtypename,invblanktypeid,invblanktype,invblanktypename,invclaimid,invclaimnumber,invotprnum,invpodnum,invsendspeedid,invsendspeedname,invsendkindid,invsendkindname,invpayplaceid,invpayplacename,invpayformid,invpayformname,invixtariffcode,invixtariffcodegdy,invannouncevalue,invavcurrencyid,invdispkindid,invrespperson,invpayplacerwcode,invpayplacerwname,
	invpayercode,--invpayerid,
	invpayername,--invpayeraccount,invpayerbank,
	invfrwsubcode,
	IF(
		length(invfrwsubcode)=11 AND substring(invfrwsubcode, 1, 3)='003',
		concat('0', invfrwsubcode),
		invfrwsubcode
	) AS invfrwsubcode_cleared,
	substring(invfrwsubcode_cleared, 3, 2) AS invfrwsubcode_indicator,
	substring(invfrwsubcode_cleared, 4, 8) AS order_id,
	--toString(toInt32OrNull(invfrwsubcode)) AS order_id,	
	--invfrwsubcode2,invloadclaim_id,invloadclaim_number,invdateexpire,invdatereceiving,invdatereceivinglocal,invfactdateaccept,invfioaccept,
	invnumber,--invuniquenumber,invgoodscashier,invgoodscashierpost,invdateready,invdatereadylocal,
	invdatedeparture,--,invrtid,invdatecustom,invdatearrive,invdatearrivelocal,invdatedelivery,invdatedeliverylocal,invdateraskrel,invdateregister,invdateregisterlocal,invdatenotification,invnotification,invnum410,invkpz,invparentid,invparentnumber,invformppsign,invtocabotageid,operstateid,operstate,operneedforecp,warning,user_trans_subcode,msg_data,changed
	carnumber
FROM 
	itrans__etran_invoice -- SELECT * FROM itrans__etran_invoice WHERE invoiceid='1661376141'
	INNER JOIN ETRAN_CAR USING invoiceid
WHERE
	--invdatedeparture >='2022-12-01' AND
	invdatedeparture >='2024-01-01' AND
	invoicestate NOT IN ('Испорчен','Сторно по отправлению отменено','Заготовка импорта, транзита')
--ORDER BY 
--	invoiceid
--LIMIT 1000000 OFFSET 0
 --AND `ETRAN_CAR.carnumber`='42007351' ORDER BY `ETRAN.invdatecreate`
--) SELECT invoiceid, count() AS q FROM ETRAN GROUP BY invoiceid HAVING q>1
--) SELECT  * FROM ETRAN
),
ETRAN AS (
SELECT
	ETRAN.*,
	contnumber
FROM
	ETRAN
	INNER JOIN ETRAN_CONT USING invoiceid
WHERE
	invdatecreate=max_invdatecreate
--) SELECT * FROM ETRAN WHERE order_id IS NULL
)
SELECT
	YOUR_EXCEL.*,
	invoiceid,
	invdatecreate,
	invfrwsubcode,
	order_id AS `подтянуто автоматически`,
	multiIf(
		order_id='', 'не определился',
		substring(order_id, 1, 2)<>'32' AND substring(order_id, 1, 2)<>'60', 'странный результат',
		Null
	) AS `проверка`
FROM
	YOUR_EXCEL
	LEFT JOIN ETRAN ON ETRAN.carnumber  = YOUR_EXCEL.`№ Вагона` AND ETRAN.contnumber  = YOUR_EXCEL.`№ Контейнера` AND ETRAN.invnumber  = YOUR_EXCEL.`Номер накладной`
"""