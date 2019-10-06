
insert into methodname
select * from 
(
select max(methodid)+1 as methodid,'ГОСТ Р 52769' as description,'ГОСТ Р 52769 (Цветность)' as textdescription from methodname a,
(select sum(case when description='ГОСТ Р 52769' then 1 else 0 end) over () as c,description from methodname a) b
where c=0 and a.description=b.description
) c
where methodid  is not null;
commit;

insert into testname 
select * from 
(
select max(testid)+1 as testid,'Цветность, градусы цветности' as description,'Цветность' as shortname from testname t,
(select sum(case when shortname ='Цветность' then 1 else 0 end) over () as c, shortname from testname a) b
where c=0 and t.shortname=b.shortname
) c
where testid is not null;

INSERT INTO METHODTESTS
select testid, methodid, 3 as testtypeid
from(
select x.testid, x.methodid, mt.testid as tid, mt.methodid as mid
from
(select t.testid, t.shortname, m.methodid, m.description
from testname t, methodname m
where t.shortname='Цветность'
and   m.DESCRIPTION='ГОСТ Р 52769') x
, methodtests mt
where x.testid = mt.testid(+)
and x.methodid = mt.methodid(+)
)
where tid is null
and mid is null;
commit;

insert into testname 
select * from 
(
select max(testid)+1 as testid,'Цветность, градусы цветности 1 определение' as description,'Цветность1опр' as shortname from testname t,
(select sum(case when shortname ='Цветность1опр' then 1 else 0 end) over () as c, shortname from testname a) b
where c=0 and t.shortname=b.shortname
) c
where testid is not null;

INSERT INTO METHODTESTS
select testid, methodid, 1 as testtypeid
from(
select x.testid, x.methodid, mt.testid as tid, mt.methodid as mid
from
(select t.testid, t.shortname, m.methodid, m.description
from testname t, methodname m
where t.shortname='Цветность1опр'
and   m.DESCRIPTION='ГОСТ Р 52769') x
, methodtests mt
where x.testid = mt.testid(+)
and x.methodid = mt.methodid(+)
)
where tid is null
and mid is null;
commit;

insert into testname 
select * from 
(
select max(testid)+1 as testid,'Цветность, градусы цветности 2 определение' as description,'Цветность2опр' as shortname from testname t,
(select sum(case when shortname ='Цветность2опр' then 1 else 0 end) over () as c, shortname from testname a) b
where c=0 and t.shortname=b.shortname
) c
where testid is not null;

INSERT INTO METHODTESTS
select testid, methodid, 2 as testtypeid
from(
select x.testid, x.methodid, mt.testid as tid, mt.methodid as mid
from
(select t.testid, t.shortname, m.methodid, m.description
from testname t, methodname m
where t.shortname='Цветность2опр'
and   m.DESCRIPTION='ГОСТ Р 52769') x
, methodtests mt
where x.testid = mt.testid(+)
and x.methodid = mt.methodid(+)
)
where tid is null
and mid is null;
commit;

INSERT INTO METHODCHARS
select distinct mc.PRODUCTID
, mt.METHODID, mt.TESTID, mc.CHARID
, mc.RANGE, mc.INACCURACY, mc.CONVERGENCE, mc.PRECISION, mc.RESULTUNIT, mc.STABREQ
, mc.NUMBERPREC, mc.POKCONVERGENCE, mc.POKLABPRECISION, mc.SISTPRECUP, mc.SISTPRECDOWN, mc.NOPRED
, mc.POKPRECISION, mc.LABID, mc.LABINACCURACYUP, mc.LABINACCURACYDOWN
from
(select
NULL as PRODUCTID, max(CHARID)+1 as CHARID,
'1 - 10' as RANGE, '0,3*X' as INACCURACY, '0,28*X' as CONVERGENCE, '0,4*X' as PRECISION, 'градусы цветности' as RESULTUNIT
, NULL as STABREQ, NULL as NUMBERPREC, NULL as POKCONVERGENCE, NULL as POKLABPRECISION, NULL as SISTPRECUP
, NULL as SISTPRECDOWN, 2 as NOPRED, NULL as POKPRECISION, NULL as LABID, NULL as LABINACCURACYUP, NULL as LABINACCURACYDOWN
from METHODCHARS) mc
,(
select tm.methodid, tm.testid, sum(case when RANGE ='1 - 10' then 1 else 0 end) over() as rg
from (
select m.methodid, t.testid
from methodname m, testname t
where m.DESCRIPTION = 'ГОСТ Р 52769'
and t.SHORTNAME = 'Цветность'
) tm, METHODCHARS mc
where mc.METHODID (+)= tm.METHODID
and mc.testid (+)= tm.testid
) mt
where mt.rg = 0;
commit;

INSERT INTO METHODCHARS
select distinct mc.PRODUCTID
, mt.METHODID, mt.TESTID, mc.CHARID
, mc.RANGE, mc.INACCURACY, mc.CONVERGENCE, mc.PRECISION, mc.RESULTUNIT, mc.STABREQ
, mc.NUMBERPREC, mc.POKCONVERGENCE, mc.POKLABPRECISION, mc.SISTPRECUP, mc.SISTPRECDOWN, mc.NOPRED
, mc.POKPRECISION, mc.LABID, mc.LABINACCURACYUP, mc.LABINACCURACYDOWN
from
(select
NULL as PRODUCTID, max(CHARID)+1 as CHARID,
'10 - 50' as RANGE, '0,2*X' as INACCURACY, '0,17*X' as CONVERGENCE, '0,24*X' as PRECISION, 'градусы цветности' as RESULTUNIT
, NULL as STABREQ, NULL as NUMBERPREC, NULL as POKCONVERGENCE, NULL as POKLABPRECISION, NULL as SISTPRECUP
, NULL as SISTPRECDOWN, 2 as NOPRED, NULL as POKPRECISION, NULL as LABID, NULL as LABINACCURACYUP, NULL as LABINACCURACYDOWN
from METHODCHARS) mc
,(
select tm.methodid, tm.testid, sum(case when RANGE ='10 - 50' then 1 else 0 end) over() as rg
from (
select m.methodid, t.testid
from methodname m, testname t
where m.DESCRIPTION = 'ГОСТ Р 52769'
and t.SHORTNAME = 'Цветность'
) tm, METHODCHARS mc
where mc.METHODID (+)= tm.METHODID
and mc.testid (+)= tm.testid
) mt
where mt.rg = 0;
commit;

INSERT INTO METHODCHARS
select distinct mc.PRODUCTID
, mt.METHODID, mt.TESTID, mc.CHARID
, mc.RANGE, mc.INACCURACY, mc.CONVERGENCE, mc.PRECISION, mc.RESULTUNIT, mc.STABREQ
, mc.NUMBERPREC, mc.POKCONVERGENCE, mc.POKLABPRECISION, mc.SISTPRECUP, mc.SISTPRECDOWN, mc.NOPRED
, mc.POKPRECISION, mc.LABID, mc.LABINACCURACYUP, mc.LABINACCURACYDOWN
from
(select
NULL as PRODUCTID, max(CHARID)+1 as CHARID,
'св. 50' as RANGE, '0,1*X' as INACCURACY, '0,1*X' as CONVERGENCE, '0,14*X' as PRECISION, 'градусы цветности' as RESULTUNIT
, NULL as STABREQ, NULL as NUMBERPREC, NULL as POKCONVERGENCE, NULL as POKLABPRECISION, NULL as SISTPRECUP
, NULL as SISTPRECDOWN, 2 as NOPRED, NULL as POKPRECISION, NULL as LABID, NULL as LABINACCURACYUP, NULL as LABINACCURACYDOWN
from METHODCHARS) mc
,(
select tm.methodid, tm.testid, sum(case when RANGE ='св. 50' then 1 else 0 end) over() as rg
from (
select m.methodid, t.testid
from methodname m, testname t
where m.DESCRIPTION = 'ГОСТ Р 52769'
and t.SHORTNAME = 'Цветность'
) tm, METHODCHARS mc
where mc.METHODID (+)= tm.METHODID
and mc.testid (+)= tm.testid
) mt
where mt.rg = 0;
commit;