

Declare @mWest datetime 
Declare @mEast char(7)
/* Get the Date of Yesterday in USA Format */
Select @mWest = DATEADD(day, 0, getdate())
/* Get the Date of Yesterday in Chinese Format */
Exec SPUtilDateToText @mWest, @mEast OUTPUT 
--print @mEast

select chOrdNo, ������=rlOrdPrio, 
�ͮĤ�=chOrdPriDay,�e��=rlOrdPriA2, ���=rlOrdPriB2,
�۶O��=case  
when chOrdPriDay < @mEast 
then rlOrdPriA2
else
	rlOrdPriB2
end, chOrdHisFlg,
chOrdOK, chOrdUnit, chOrdDct, 
* from GenOrdBasicTbl (nolock)
where rtrim(isnull(chOrdOK,'')) <>'0'
and rtrim(isnull(chOrdUnit,'')) not in('��','�`','��','��')
and rtrim(isnull(chOrdDct,'')) in ('12', '17','75')
and rlOrdPrio > case  
when chOrdPriDay < @mEast 
then rlOrdPriA2
else
	rlOrdPriB2
end *0.9


--and rtrim(isnull(chOrdHisFlg,'')) in ('A', 'H', 'I')



