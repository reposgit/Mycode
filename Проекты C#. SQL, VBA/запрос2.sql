(select max(sampledate) as sampledate,  '' as head,lab, productdescription, a.orderattribute, to_char(count(distinct fc)) as RES,'к-во анализов' as text , a.testname as testname,
replace(sampleplace,level1||',','') as AGREGAT, GetQuartName(max(a.sampledate)) as qr, 'За '||GetQuartName(max(a.sampledate))||' '||to_char(a.sampledate, 'yyyy')||' года' as quartyear
 from allslresult a
 where %SAMPLEDATE%
and productdescription like '%Масло%'  and orderattribute not like '%ГОСТ 6307_pH%'
group by a.lab,productdescription, replace(sampleplace,level1||',',''), orderattribute,testname, sampledate
UNION ALL
select resvz.sampledate, '' as head, a.lab,a.productdescription,a.orderattribute,res,'посл. знач.,'||resultunit as text, testname,replace(sampleplace,level1||',','') as AGREGAT, 
GetQuartName(a.sampledate) as qr, 'За '||GetQuartName(a.sampledate)||' '||to_char(a.sampledate, 'yyyy')||' года' as quartyear
 from allslresult a ,  
 (select lab,productdescription,orderattribute,replace(sampleplace,level1||',','') as AGREGAT,max(sampledate) as sampledate from allslresult 
    where %SAMPLEDATE%
    and productdescription like '%Масло%' 
   and orderattribute not like '%ГОСТ 1547_НаличиеВоды%' and orderattribute not like '%ГОСТ 6307_pH%' and orderattribute not like '%ГОСТ 18995.1_ПлотностьКГМ%' and orderattribute not like '%ГОСТ 18995.1_ПлотностьГ%'
    group by lab,productdescription, orderattribute,testname, sampledate, replace(sampleplace,level1||',','')) resvz
 where A.sampledate = resvz.sampledate
  and a.productdescription=resvz.productdescription
  and replace(a.sampleplace,a.level1||',','') = agregat
  and a.productdescription like '%Масло%'
  and A.ORDERATTRIBUTE = resvz.orderattribute
  and a.%SAMPLEDATE%)