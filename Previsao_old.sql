use colibri8

DECLARE @BaseDate smalldatetime
DECLARE @BaseTime time
DECLARE @Data as datetime
DECLARE @total as money 

SET @BaseDate = '2017-01-02'
SET @BaseTime = CAST(getdate() as time) --'12:45'
SET @Data = GETDATE()

DECLARE @ticket int
DECLARE @media_antes int
DECLARE @media_depois int
DECLARE @min_antes int
DECLARE @max_antes int
DECLARE @min_depois int
DECLARE @max_depois int
DECLARE @valor money
DECLARE @media_valor money
DECLARE @min_valor money
DECLARE @max_valor money
DECLARE @ticket_medio money
DECLARE @desconto money
DECLARE @desc_globo money
DECLARE @desc_outros money
DECLARE @n_globo int
DECLARE @media_globo int
DECLARE @n_desc_globo int
DECLARE @n_desc_outros int
DECLARE @bebida int
DECLARE @comida float(24)
DECLARE @sobremesa int
DECLARE @bomboniere int

declare @tblFinal as table (
  fichas int,
  media_antes int,
  media_depois int,
  valor numeric(15,2))

declare @tbl2 as table (
  data date,
  soma int,
  valor numeric(15,2))

declare @tblGlobo as table (
  c int)
------------------------------------------------------------------------------------

select @total = sum(m.vl_desconto) 
from vw_movimento m with (nolock) 
where vl_desconto > 0
  and cancelada = 0  
  and dt_contabil = FORMAT(@Data,'yyyyMMdd')

declare @tbl as table (
  mododevenda_id int,
  mododevenda varchar(10),
  qtd_pessoas int,
  total money)

----------------------------------------------------------------------------------

insert into @tbl
select
  ov.mododevenda_id,
  upper(min(md.nome)),
  qtd_pessoas = ov.nu_pessoas,
  total = sum(isnull(m.vl_recebido, 0))
from operacaogeral o with (nolock)
join operacaodevendageral ov with (nolock) on o.operacao_id = ov.operacao_id
join movimentocaixageral m with (nolock) on m.operacao_id = o.operacao_id
join mododevenda md with (nolock) on ov.mododevenda_id = md.id
where isnull(m.cancelado, 0) = 0
  and o.tipo = 'venda'
  and ov.encerrada = 1
  and o.cancelada = 0
  
  and o.dt_contabil = FORMAT(@Data,'yyyyMMdd')

group by o.operacao_id, ov.mododevenda_id, ov.nu_pessoas

insert into @tbl
select
  ov.mododevenda_id,
  upper(min(md.nome)),
  qtd_pessoas = ov.nu_pessoas,
  total = sum(isnull(m.vl_recebido, 0))
from operacao o with (nolock)
join operacaodevenda ov with (nolock) on o.operacao_id = ov.operacao_id
join movimentocaixa m with (nolock) on m.operacao_id = o.operacao_id
join mododevenda md with (nolock) on ov.mododevenda_id = md.id
where isnull(m.cancelado, 0) = 0
  and o.tipo = 'venda'
  and ov.encerrada = 1
  and o.cancelada = 0
  
  and o.dt_contabil = FORMAT(@Data,'yyyyMMdd')
group by o.operacao_id, ov.mododevenda_id, ov.nu_pessoas


--select	@ticket = a.fichas_abertas + a.fichas_pagas, @valor = a.ticket_medio * (a.fichas_abertas + a.fichas_pagas), @desconto = c.desconto, @ticket_medio = (b.total - c.desconto)/(a.fichas_abertas + a.fichas_pagas)
select	@ticket = a.fichas_abertas + a.fichas_pagas, @valor = b.total - c.desconto, @desconto = c.desconto, @ticket_medio = (b.total - c.desconto)/(a.fichas_abertas + a.fichas_pagas)
from
(select
  fichas_abertas = (SELECT count(
      [status])
  FROM [colibri8].[dbo].[ficha]
  where status = 'consumindo'),
  fichas_pagas = sum(qtd_pessoas),
  total = sum(total),
  tipo = case
    when mododevenda_id in (3, 4) then 'pessoas'
    else 'atendimentos'
  end,
  ticket_medio = round(sum(total) / sum(qtd_pessoas), 2,1)
from @tbl
group by mododevenda_id,mododevenda) a,

(select
  total = sum(i.vl_tot)
from vw_itemvendaunificada i with (nolock)
join
(
  select total = sum(i.vl_tot)
  from vw_itemvendaunificada i with (nolock)
  where i.cancelado = 0
    and i.transferido = 0
    
  and i.dt_contabil = FORMAT(@Data,'yyyyMMdd')
) x on 1=1
where i.cancelado = 0
  and i.transferido = 0
  
  and i.dt_contabil = FORMAT(@Data,'yyyyMMdd')) b,

(select desconto = isnull((select top 1 @total
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  
  and dt_contabil = FORMAT(@Data,'yyyyMMdd')
group by nm_desconto), 0)) c

--------------------------------------------------------------------------------------------
insert into @tbl2
  SELECT CONVERT(date, [dt_abertura]), COUNT(*), SUM(vl_subtotal_para_servico)
    FROM [colibri8].[dbo].[headervendageral]
  WHERE CONVERT(date, [dt_abertura]) > FORMAT(@BaseDate,'yyyy-MM-dd')
  and CAST([dt_abertura] as time) < @BaseTime
  and DATEPART(DW, dt_abertura) # 7
  GROUP BY CONVERT(date, [dt_abertura])

SELECT @media_antes = AVG(soma), @min_antes = MIN(soma), @max_antes = MAX(soma) from @tbl2

DELETE FROM @tbl2

insert into @tbl2
  SELECT CONVERT(date, [dt_abertura]), COUNT(*), SUM(vl_subtotal_para_servico)
    FROM [colibri8].[dbo].[headervendageral]
  WHERE CONVERT(date, [dt_abertura]) > FORMAT(@BaseDate,'yyyy-MM-dd')
  and CAST([dt_abertura] as time) > @BaseTime
  and DATEPART(DW, dt_abertura) # 7
  GROUP BY CONVERT(date, [dt_abertura])

SELECT @media_depois = AVG(soma), @min_depois = MIN(soma), @max_depois = MAX(soma) from @tbl2


declare @tbl3 as table (
  data date,
  valor numeric(15,2))

insert into @tbl3
	SELECT FORMAT(dt_abertura,'yyyy-MM-dd'), SUM([vl_subtotal_para_servico])
	FROM [colibri8].[dbo].[headervendageral]
	WHERE CONVERT(date, [dt_abertura]) > FORMAT(@BaseDate,'yyyy-MM-dd')
	and CAST([dt_abertura] as time) < @BaseTime
	and DATEPART(DW, dt_abertura) # 7
	and [cancelado] = 0
	GROUP BY FORMAT(dt_abertura,'yyyy-MM-dd')
	ORDER BY FORMAT(dt_abertura,'yyyy-MM-dd')

--select * from @tbl3 order by data

declare @tbl4 as table (
  data date,
  valor numeric(15,2))

declare
  @totalM as money 

select @totalM = sum(m.vl_desconto) 
from vw_movimento m with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  
  and dt_contabil > @BaseDate

insert into @tbl4
select 
CONVERT(date, [dt_contabil]),
  [Valor desconto] = sum(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and CONVERT(date, [dt_contabil]) > FORMAT(@BaseDate,'yyyy-MM-dd')
  and DATEPART(DW, dt_contabil) # 7
group by CONVERT(date, [dt_contabil])
order by CONVERT(date, [dt_contabil])

select @media_valor = CONVERT(money,AVG(a.valor - b.valor)),
		@min_valor = CONVERT(money,MIN(a.valor - b.valor)),
		@max_valor = CONVERT(money,MAX(a.valor - b.valor))
from @tbl3 a join @tbl4 b ON a.data = b.data

--------- DESCONTOS ---------------

select 
  @desc_globo = sum(vl_desconto),
  @n_desc_globo = count(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and dt_contabil = FORMAT(@Data,'yyyy-MM-dd')
  and nm_desconto = 'TV GLOBO'

select 
  @n_globo = count(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and dt_contabil = FORMAT(@Data,'yyyy-MM-dd')
  and nm_desconto LIKE '%GLOBO%'

insert into @tblGlobo
select 
  count(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and CONVERT(date, [dt_contabil]) > FORMAT(@BaseDate,'yyyy-MM-dd')
  and CAST([dt_recebimento] as time) < @BaseTime
  and DATEPART(DW, dt_contabil) # 7
  and nm_desconto LIKE '%GLOBO%'
group by dt_contabil

select @media_globo = AVG(c) from @tblGlobo

select 
  @desc_outros = sum(vl_desconto),
  @n_desc_outros = count(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and dt_contabil = FORMAT(@Data,'yyyy-MM-dd')
  and nm_desconto <> 'TV GLOBO'
-------------------------------------

----- VENDAS ------------------------
select
  @bebida = sum(i.qtd)
from vw_itemvendaunificada i with (nolock)
where i.cancelado = 0
  and i.transferido = 0  
  and i.grupo_id = 9
  and i.dt_contabil = FORMAT(@Data,'yyyy-MM-dd')
group by
  i.grupo_id

select
  @comida = sum(i.qtd)
from vw_itemvendaunificada i with (nolock)
where i.cancelado = 0
  and i.transferido = 0
  and i.grupo_id = 12
  and i.dt_contabil = FORMAT(@Data,'yyyy-MM-dd')
group by
  i.grupo_id

select
  @sobremesa = sum(i.qtd)
from vw_itemvendaunificada i with (nolock)
where i.cancelado = 0
  and i.transferido = 0
  and i.grupo_id = 4
  and i.dt_contabil = FORMAT(@Data,'yyyy-MM-dd')
group by
  i.grupo_id

select
  @bomboniere = sum(i.qtd)
from vw_itemvendaunificada i with (nolock)
where i.cancelado = 0
  and i.transferido = 0
  and i.grupo_id = 11
  and i.dt_contabil = FORMAT(@Data,'yyyy-MM-dd')
group by
  i.grupo_id
------------------------------------

select  ISNULL(@ticket, 0) as tickets,
		ISNULL(@ticket + @media_depois, 0) as proj,
		ISNULL(@ticket - @media_antes, 0) as dif_media,
		ISNULL(@ticket - @max_antes, 0) as dif_max,
		ISNULL(@media_antes, 0) as media_antes,
		ISNULL(@min_antes, 0) as min_antes,
		ISNULL(@max_antes, 0) as max_antes,
		ISNULL(@media_depois, 0) as media_depois, 
		ISNULL(@min_depois, 0) as min_depois,
		ISNULL(@max_depois, 0) as max_depois,
		ISNULL(@valor, 0) as valor, 
		ISNULL(@media_valor, 0) as media_valor, 
		ISNULL(@min_valor, 0) as min_valor, 
		ISNULL(@max_valor, 0) as max_valor,
		ISNULL(@ticket_medio, 0) as ticket_medio,
		ISNULL(@desconto, 0) as desconto,
		ISNULL(@desc_globo, 0) as desc_globo,
		ISNULL(@desc_outros, 0) as desc_outros,
		ISNULL(@n_globo, 0) as n_globo,
		ISNULL(@n_desc_globo, 0) as n_desc_globo,
		ISNULL(@n_desc_outros, 0) as n_desc_outros,
		ISNULL(@media_globo, 0) as media_globo,
		ISNULL(@bebida, 0) as bebida,
		ISNULL(@comida, 0) as comida,
		ISNULL(@sobremesa, 0) as sobremesa,
		ISNULL(@bomboniere, 0) as bomboniere