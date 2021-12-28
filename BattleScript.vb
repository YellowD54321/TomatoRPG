
Dim monsterState(10) as Variant
Dim playerState(10) as Variant

Sub 戰鬥試算巨集()
	'
	' 戰鬥試算巨集
	'
	Dim sheetName As String
	Dim resultSheetName As String
	
	Dim monsterName As String
	Dim monsterMaxHp As Single
	Dim monsterAttack As Single
	Dim monsterDefence As Single
	Dim monsterFlee As Single
	Dim monsterCrit As Single
	Dim monsterAttackDistance As Single
	Dim monsterDamageReduceFomula As String
	Dim monsterDamageReduce As Single
	Dim monsterStateStartRow As String
	Dim monsterStateStartColumn As String
	Dim monsterStateNextColumn As String
	Dim monsterDefenceCoordinate(2) As Variant
	Dim equipName As String
	Dim equipMaxHp As Single
	Dim equipAttack As Single
	Dim equipDefence As Single
	Dim equipFlee As Single
	Dim equipCrit As Single
	Dim equipAttackDistance As Single
	Dim equipIncreaseHp As Single
	Dim equipDamageReduceFomula As String
	Dim equipDamageReduce As Single
	Dim equipStateStartRow As String
	Dim equipStateStartColumn As String
	Dim equipStateNextColumn As String
	Dim equipDefenceCoordinate(2) As Variant
	
	Dim battleHit as single
	Dim battleCrit as single
	Dim battleDamageResult as Variant
	Dim battleDamage as single
	Dim battleDamageRandomPercent as single
	Dim battleTurn as integer
	
	Dim firstAttacker as Variant
	Dim secondAttacker as Variant
	Dim firstAttackerWinTimes as integer
	Dim secondAttackerWinTimes as integer
	
	Dim recorderCoordinationRow as String
	Dim recorderCoordinationColumn as String
	Dim titleCoordinationRow as String
	Dim titleCoordinationColumn as String
	
	
	resultSheetName = "戰鬥試算"
	sheetName = "主頁面"
	monsterStateStartRow = "4"
	monsterStateStartColumn = "C"
	monsterName = Sheets(sheetName).Cells(monsterStateStartRow, monsterStateStartColumn).Value

	monsterStateNextColumn = getNextColumn(sheetName, monsterStateStartColumn)
	monsterMaxHp = getNextColumnValue(sheetName, monsterStateStartRow, monsterStateNextColumn)

	monsterStateNextColumn = getNextColumn(sheetName, monsterStateNextColumn)
	monsterAttack = getNextColumnValue(sheetName, monsterStateStartRow, monsterStateNextColumn)

	monsterStateNextColumn = getNextColumn(sheetName, monsterStateNextColumn)
	monsterDefence = getNextColumnValue(sheetName, monsterStateStartRow, monsterStateNextColumn)
	monsterDefenceCoordinate(0) = monsterStateStartRow
	monsterDefenceCoordinate(1) = monsterStateNextColumn

	monsterStateNextColumn = getNextColumn(sheetName, monsterStateNextColumn)
	monsterFlee = getNextColumnValue(sheetName, monsterStateStartRow, monsterStateNextColumn)

	monsterStateNextColumn = getNextColumn(sheetName, monsterStateNextColumn)
	monsterCrit = getNextColumnValue(sheetName, monsterStateStartRow, monsterStateNextColumn)

	monsterStateNextColumn = getNextColumn(sheetName, monsterStateNextColumn)
	monsterAttackDistance = getNextColumnValue(sheetName, monsterStateStartRow, monsterStateNextColumn)

	monsterDamageReduceFomula = "=(" + CStr(monsterDefenceCoordinate(1)) + CStr(monsterDefenceCoordinate(0)) + "/10)^(1/2)*7 * 1/100"
	
	Sheets(sheetName).Cells("4", "K") = monsterDamageReduceFomula
	monsterDamageReduce = Round(Sheets(sheetName).Cells("4", "K"),4)
	
	equipStateStartRow = "5"
	equipStateStartColumn = "C"
	equipName = Sheets(sheetName).Cells(equipStateStartRow, equipStateStartColumn).Value

	equipStateNextColumn = getNextColumn(sheetName, equipStateStartColumn)
	equipMaxHp = getNextColumnValue(sheetName, equipStateStartRow, equipStateNextColumn)

	equipStateNextColumn = getNextColumn(sheetName, equipStateNextColumn)
	equipAttack = getNextColumnValue(sheetName, equipStateStartRow, equipStateNextColumn)

	equipStateNextColumn = getNextColumn(sheetName, equipStateNextColumn)
	equipDefence = getNextColumnValue(sheetName, equipStateStartRow, equipStateNextColumn)
	equipDefenceCoordinate(0) = equipStateStartRow
	equipDefenceCoordinate(1) = equipStateNextColumn

	equipStateNextColumn = getNextColumn(sheetName, equipStateNextColumn)
	equipFlee = getNextColumnValue(sheetName, equipStateStartRow, equipStateNextColumn)

	equipStateNextColumn = getNextColumn(sheetName, equipStateNextColumn)
	equipCrit = getNextColumnValue(sheetName, equipStateStartRow, equipStateNextColumn)

	equipStateNextColumn = getNextColumn(sheetName, equipStateNextColumn)
	equipAttackDistance = getNextColumnValue(sheetName, equipStateStartRow, equipStateNextColumn)
	
	equipStateNextColumn = getNextColumn(sheetName, equipStateNextColumn)
	equipIncreaseHp = getNextColumnValue(sheetName, equipStateStartRow, equipStateNextColumn)

	equipDamageReduceFomula = "=(" + CStr(equipDefenceCoordinate(1)) + CStr(equipDefenceCoordinate(0)) + "/10)^(1/2)*7 * 1/100"

	Sheets(sheetName).Cells("5", "K") = equipDamageReduceFomula
	equipDamageReduce = Round(Sheets(sheetName).Cells("5", "K"),4)
	
	monsterState(0) = monsterName
	monsterState(1) = CSng(monsterMaxHp)
	monsterState(2) = CSng(monsterMaxHp) ' current hp
	monsterState(3) = CSng(monsterAttack)
	monsterState(4) = CSng(monsterDefence)
	monsterState(5) = CSng(monsterFlee)
	monsterState(6) = CSng(monsterCrit)
	monsterState(7) = CSng(monsterAttackDistance)
	monsterState(8) = 0 ' increase hp 
	monsterState(9) = CSng(monsterDamageReduce)
	monsterState(10) = false ' isDead
	
	playerState(0) = equipName
	playerState(1) = CSng(equipMaxHp)
	playerState(2) = CSng(equipMaxHp) ' current hp
	playerState(3) = CSng(equipAttack)
	playerState(4) = CSng(equipDefence)
	playerState(5) = CSng(equipFlee)
	playerState(6) = CSng(equipCrit)
	playerState(7) = CSng(equipAttackDistance)
	playerState(8) = CSng(equipIncreaseHp)
	playerState(9) = CSng(equipDamageReduce)
	playerState(10) = false ' isDead
	
	' show battle result sheet
	Sheets(resultSheetName).Select
	
	' clear sheet
	Sheets(resultSheetName).Cells.ClearContents
	
	' set first attacker and second attacker
	if monsterAttackDistance > equipAttackDistance Then
		firstAttacker = monsterState
		secondAttacker = playerState
	elseif monsterAttackDistance < equipAttackDistance Then
		firstAttacker = playerState
		secondAttacker = monsterState
	Else
		Dim rnadomDice as single
		rnadomDice = Int((100 - 0 + 1) * Rnd() + 0)
		if rnadomDice >= 50 Then
			firstAttacker = "monster"
			secondAttacker = "player"
		Else
			firstAttacker = "player"
			secondAttacker = "monster"
		end if
	end if
	
	Sheets(resultSheetName).Cells("2", "E") = "先攻"
	Sheets(resultSheetName).Cells("3", "E") = "後攻"
	Sheets(resultSheetName).Cells("2", "F") = firstAttacker(0)
	Sheets(resultSheetName).Cells("3", "F") = secondAttacker(0)
	
	' battle start
	titleCoordinationRow = 5
	' titleCoordinationColumn = "A"
	
	dim running as integer
	dim runTimes as integer
	runTimes = Sheets(sheetName).Cells("7", "D").Value
	firstAttackerWinTimes = 0
	secondAttackerWinTimes = 0
	for running = 1 to runTimes
		titleCoordinationColumn = "A"
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "第" + CStr(running) + "場戰鬥"
		' reset state
		firstAttacker(2) = firstAttacker(1)
		firstAttacker(10) = false
		secondAttacker(2) = secondAttacker(1)
		secondAttacker(10) = false
		'write title
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "回合數"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "後攻當前血量"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "先攻傷害"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "後攻剩餘血量"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "是否暴擊"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "是否命中"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "傷害浮動%"
		titleCoordinationColumn = "J"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "先攻當前血量"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "後攻傷害"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "先攻剩餘血量"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "是否暴擊"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "是否命中"
		titleCoordinationColumn = getNextColumn(sheetName, titleCoordinationColumn)
		Sheets(resultSheetName).Cells(titleCoordinationRow, titleCoordinationColumn) = "傷害浮動%"
		
		battleTurn = 0
		
		do while battleTurn <= 200
			battleTurn = battleTurn + 1
			firstAttacker(10) = isDead(firstAttacker)
			secondAttacker(10) = isDead(secondAttacker)
			if firstAttacker(10) Or secondAttacker(10) Then
				' MsgBox("回合開始，有人血量歸零，戰鬥結束")
				Exit Do
			end if
			' record turn number
			recorderCoordinationRow = titleCoordinationRow + battleTurn
			recorderCoordinationColumn = "B"
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleTurn
			' first attack
			' damage calculate
			battleHit = battleIsFlee(0, secondAttacker(5))
			battleCrit = battleIsCrit(firstAttacker(6))
			battleDamageResult = damageCalculate(firstAttacker(3), secondAttacker(9))
			battleDamage = Round(battleDamageResult(0) * battleCrit * battleHit)
			battleDamageRandomPercent = battleDamageResult(1)
			' record first attack battle record
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = secondAttacker(2)
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleDamage
			' damage on second attcker
			secondAttacker(2) = secondAttacker(2) - battleDamage
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = secondAttacker(2)
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleCrit
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleHit
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleDamageRandomPercent
			
			' check if second attacker is dead
			secondAttacker(10) = isDead(secondAttacker)
			if secondAttacker(10) Then
				firstAttackerWinTimes = firstAttackerWinTimes + 1
				Sheets(resultSheetName).Cells(titleCoordinationRow-1, "I") = "先攻勝利"
				' MsgBox("後攻者血量歸零，戰鬥結束")
				Exit Do
			end if
		
			' second attack
			recorderCoordinationColumn = "J"
			' damage calculate
			battleHit = battleIsFlee(0, firstAttacker(5))
			battleCrit = battleIsCrit(secondAttacker(6))
			battleDamageResult = damageCalculate(secondAttacker(3), firstAttacker(9))
			battleDamage = Round(battleDamageResult(0) * battleCrit * battleHit)
			battleDamageRandomPercent = battleDamageResult(1)
			' record second attack battle record
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = firstAttacker(2)
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleDamage
			' damage on second attcker
			firstAttacker(2) = firstAttacker(2) - battleDamage
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = firstAttacker(2)
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleCrit
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleHit
			recorderCoordinationColumn = getNextColumn(sheetName, recorderCoordinationColumn)
			Sheets(resultSheetName).Cells(recorderCoordinationRow, recorderCoordinationColumn) = battleDamageRandomPercent
		
			' check if first attacker is dead
			firstAttacker(10) = isDead(firstAttacker)
			if firstAttacker(10) Then
				secondAttackerWinTimes = secondAttackerWinTimes + 1
				Sheets(resultSheetName).Cells(titleCoordinationRow-1, "I") = "後攻勝利"
				' MsgBox("先攻者血量歸零，戰鬥結束")
				Exit Do
			end if
		Loop
	
		titleCoordinationRow = recorderCoordinationRow + 2
		titleCoordinationColumn = "B"
	next running
	' record win rate
	Sheets(resultSheetName).Cells("1", "F") = "總戰鬥場數：" + CStr(runTimes)
	Sheets(resultSheetName).Cells("1", "G") = "勝率"
	Sheets(resultSheetName).Cells("2", "G") = firstAttackerWinTimes / (firstAttackerWinTimes + secondAttackerWinTimes)
	Sheets(resultSheetName).Cells("3", "G") = secondAttackerWinTimes / (firstAttackerWinTimes + secondAttackerWinTimes)
	
	Sheets(resultSheetName).Columns("A:P").AutoFit
	Sheets(resultSheetName).Columns("A:P").HorizontalAlignment = xlCenter
End Sub

Function getNextColumnValue(sheetName as String, nextRow As String, nextColumn As String) As String
	Dim monsterStateNextRow As String
	getNextColumnValue =  Sheets(sheetName).Cells(nextRow, nextColumn).Value
End Function

Function getNextColumn(sheetName as String, lastColumn As String) As String
	Dim monsterStateNextColumn As String
	Dim monsterStateNextColumnAsNumber As Single
	monsterStateNextColumnAsNumber = Range(lastColumn & 1).Column + 1
	monsterStateNextColumn = Split(Cells(1, monsterStateNextColumnAsNumber).Address, "$")(1)
	getNextColumn = monsterStateNextColumn
End Function

function damageCalculate(actorAttack as Variant, defenderDR as Variant) As Variant
	dim damageResult(1) as Variant
	dim finalDamage as single
	Dim rnadomDice as single
	dim damageRangeMax as integer
	dim damageRangeMin as integer
	damageRangeMax = 100
	damageRangeMin = 85
	rnadomDice = Int((damageRangeMax - damageRangeMin + 1) * Rnd() + damageRangeMin)
	finalDamage = actorAttack * (1 - defenderDR) * (rnadomDice / 100)
	damageResult(0) = finalDamage
	damageResult(1) = rnadomDice
	damageCalculate = damageResult
End Function

function accurationCalculate(actorAccuration as Variant, defenderFlee as Variant) as single
	accurationCalculate = 1 + (actorAccuration / 100 / 2) - (defenderFlee / 100)
end function

function battleIsCrit(actorCrit as Variant) as single
	Dim rnadomDice as single
	Dim randomMax as integer
	Dim randomMin as integer
	Dim crit as single
	randomMax = 10000
	randomMin = 0
	crit = actorCrit * 100
	rnadomDice = Int((randomMax - randomMin + 1) * Rnd() + randomMin)
	if rnadomDice <= crit Then
		battleIsCrit = 1.5
	Else
		battleIsCrit = 1
	end if
end function

function battleIsFlee(actorAccuration as Variant, defenderFlee as Variant) as single
	Dim accuRate as single
	Dim rnadomDice as single
	Dim randomMax as integer
	Dim randomMin as integer
	accuRate = accurationCalculate(actorAccuration, defenderFlee) * 10000
	randomMax = 10000
	randomMin = 0
	rnadomDice = Int((randomMax - randomMin + 1) * Rnd() + randomMin)
	if rnadomDice <= accuRate Then
		battleIsFlee = 1
	Else
		battleIsFlee = 0
	end if
end function

function isDead(target as Variant) as boolean
	' check current hp
	if target(2) <= 0 Then
		isDead = true
	Else
		isDead = False
	end if
end function