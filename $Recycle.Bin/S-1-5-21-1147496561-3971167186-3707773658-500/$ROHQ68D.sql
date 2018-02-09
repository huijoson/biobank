

Declare @mWest datetime 
Declare @mEast char(7)
/* Get the Date of Yesterday in USA Format */
Select @mWest = DATEADD(day, 0, getdate())
/* Get the Date of Yesterday in Chinese Format */
Exec SPUtilDateToText @mWest, @mEast OUTPUT 
--print @mEast

select chOrdNo, Θセ基=rlOrdPrio, 
ネら=chOrdPriDay,玡基=rlOrdPriA2, 基=rlOrdPriB2,
禣基=case  
when chOrdPriDay < @mEast 
then rlOrdPriA2
else
	rlOrdPriB2
end, chOrdHisFlg,
chOrdOK, chOrdUnit, chOrdDct, 
* from GenOrdBasicTbl (nolock)
where rtrim(isnull(chOrdOK,'')) <>'0'
and rtrim(isnull(chOrdUnit,'')) not in('Ω','','そ','だ')
and rtrim(isnull(chOrdDct,'')) in ('12', '17','75')
and rlOrdPrio > case  
when chOrdPriDay < @mEast 
then rlOrdPriA2
else
	rlOrdPriB2
end *0.9


--and rtrim(isnull(chOrdHisFlg,'')) in ('A', 'H', 'I')



