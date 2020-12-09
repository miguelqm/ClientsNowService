use colibri8

DECLARE @BaseDate smalldatetime
DECLARE @BaseTime time
DECLARE @StartTime time
DECLARE @EndTime time
DECLARE @Data as datetime
DECLARE @total as numeric(10,2) 

SET @BaseDate = GETDATE()-280
SET @StartTime = '11:30'--'#STARTTIME#'
SET @EndTime = '16:59'--'#ENDTIME#'
SET @BaseTime = CAST(getdate() as time) --'12:45'
SET @Data = CAST(GETDATE() as date)

DECLARE @isNovoPeriodo char
DECLARE @ticket int
DECLARE @media_antes int
DECLARE @media_depois int
DECLARE @min_antes int
DECLARE @max_antes int
DECLARE @min_depois int
DECLARE @max_depois int
DECLARE @valor numeric(10,2)
DECLARE @media_valor numeric(10,2)
DECLARE @min_valor numeric(10,2)
DECLARE @max_valor numeric(10,2)
DECLARE @ticket_medio numeric(10,2)
DECLARE @desconto numeric(10,2)
DECLARE @desc_globo numeric(10,2)
DECLARE @desc_outros numeric(10,2)
DECLARE @n_globo int
DECLARE @media_globo int
DECLARE @n_desc_globo int
DECLARE @n_desc_outros int
DECLARE @bebida int
DECLARE @comida numeric(10,2)
DECLARE @sobremesa int
DECLARE @bomboniere int
DECLARE @total_mes numeric(10,2)
DECLARE @media_comida numeric(6,2)
DECLARE @comida_por_pessoa numeric(6,3)
DECLARE @media_quilo numeric(6,3)
DECLARE @fichas_abertas int

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
------------------------------------------------------------------------------------FORMAT(dt_recebimento as date)

select @isNovoPeriodo = IIF(cont = rec, '', '*') 
from (
select top 1 cont = CAST(dt_contabil as date), rec = CAST(dt_recebimento as date)
from vw_movimento m with (nolock) 
where vl_desconto > 0
  and cancelada = 0  
  and CAST(dt_recebimento as date) = @Data
  and CAST(dt_recebimento as time) > @StartTime
  and CAST(dt_recebimento as time) < @EndTime
order by dt_recebimento desc) v

select @total = sum(m.vl_desconto)
from vw_movimento m with (nolock) 
where vl_desconto > 0
  and cancelada = 0  
  and CAST(dt_recebimento as date) = @Data
  and CAST(dt_recebimento as time) > @StartTime
  and CAST(dt_recebimento as time) < @EndTime

declare @tbl as table (
  mododevenda_id int,
  mododevenda varchar(10),
  qtd_pessoas int,
  total numeric(10,2))

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
  
  and CAST(o.dt_alt as date) = @Data

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
  and CAST(o.dt_alt as time) > @StartTime
  and CAST(o.dt_alt as time) < @EndTime
  and CAST(o.dt_alt as date) = @Data
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
    and CAST(i.dt_lanc as time) > @StartTime
	and CAST(i.dt_lanc as time) < @EndTime
  and CAST(i.dt_lanc as date) = @Data
) x on 1=1
where i.cancelado = 0
  and i.transferido = 0
  and CAST(i.dt_lanc as time) > @StartTime
  and CAST(i.dt_lanc as time) < @EndTime
  and CAST(i.dt_lanc as date) = @Data) b,

(select desconto = isnull((select top 1 @total
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and CAST(dt_recebimento as time) > @StartTime
  and CAST(dt_recebimento as time) < @EndTime
  and CAST(dt_recebimento as date) = @Data
group by nm_desconto), 0)) c

--------------------------------------------------------------------------------------------
insert into @tbl2
  SELECT CONVERT(date, [dt_abertura]), COUNT(*), SUM(vl_subtotal_para_servico)
    FROM [colibri8].[dbo].[headervendageral]
  WHERE CONVERT(date, [dt_abertura]) > CAST(@BaseDate as date)
  and CAST([dt_abertura] as time) < @BaseTime
  and CAST([dt_abertura] as time) > @StartTime
  and CAST([dt_abertura] as time) < @EndTime
  and DATEPART(DW, dt_abertura) = DATEPART(DW, @Data)
  GROUP BY CONVERT(date, [dt_abertura])

SELECT @media_antes = AVG(soma), @min_antes = MIN(soma), @max_antes = MAX(soma) from @tbl2

DELETE FROM @tbl2

insert into @tbl2
  SELECT CONVERT(date, [dt_abertura]), COUNT(*), SUM(vl_subtotal_para_servico)
    FROM [colibri8].[dbo].[headervendageral]
  WHERE CONVERT(date, [dt_abertura]) > CAST(@BaseDate as date)
  and CAST([dt_abertura] as time) > @BaseTime
  and CAST([dt_abertura] as time) > @StartTime
  and CAST([dt_abertura] as time) < @EndTime
  and DATEPART(DW, dt_abertura) = DATEPART(DW, @Data)
  GROUP BY CONVERT(date, [dt_abertura])

SELECT @media_depois = AVG(soma), @min_depois = MIN(soma), @max_depois = MAX(soma) from @tbl2


declare @tbl3 as table (
  data date,
  valor numeric(15,2))

insert into @tbl3
	SELECT CAST(dt_abertura as date), SUM([vl_subtotal_para_servico])
	FROM [colibri8].[dbo].[headervendageral]
	WHERE CONVERT(date, [dt_abertura]) > CAST(@BaseDate as date)
	and CAST([dt_abertura] as time) < @BaseTime
	and CAST([dt_abertura] as time) > @StartTime
	and CAST([dt_abertura] as time) < @EndTime
	and DATEPART(DW, dt_abertura) = DATEPART(DW, @Data)
	and [cancelado] = 0
	GROUP BY CAST(dt_abertura as date)
	ORDER BY CAST(dt_abertura as date)

--select * from @tbl3 order by data

declare @tbl4 as table (
  data date,
  valor numeric(15,2))

declare
  @totalM as numeric(10,2) 

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
  and CONVERT(date, [dt_contabil]) > CAST(@BaseDate as date)
  and CAST(dt_recebimento as time) > @StartTime
  and CAST(dt_recebimento as time) < @EndTime
  and DATEPART(DW, dt_contabil) = DATEPART(DW, @Data)
group by CONVERT(date, [dt_contabil])
order by CONVERT(date, [dt_contabil])

select @media_valor = CONVERT(numeric(10,2),AVG(a.valor - b.valor)),
		@min_valor = CONVERT(numeric(10,2),MIN(a.valor - b.valor)),
		@max_valor = CONVERT(numeric(10,2),MAX(a.valor - b.valor))
from @tbl3 a join @tbl4 b ON a.data = b.data

--------- DESCONTOS ---------------

select 
  @desc_globo = sum(vl_desconto),
  @n_desc_globo = count(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and CAST(dt_recebimento as date) = @Data
  and CAST(dt_recebimento as time) > @StartTime
  and CAST(dt_recebimento as time) < @EndTime
  and nm_desconto = 'TV GLOBO'

select 
  @n_globo = count(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and CAST(dt_recebimento as date) = @Data
  and CAST(dt_recebimento as time) > @StartTime
  and CAST(dt_recebimento as time) < @EndTime
  and nm_desconto LIKE '%GLOBO%'

insert into @tblGlobo
select 
  count(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and CONVERT(date, [dt_contabil]) > CAST(@BaseDate as date)
  and CAST([dt_recebimento] as time) < @BaseTime
  and CAST(dt_recebimento as time) > @StartTime
  and CAST(dt_recebimento as time) < @EndTime
  and DATEPART(DW, dt_contabil) = DATEPART(DW, @Data)
  and nm_desconto LIKE '%GLOBO%'
group by dt_contabil

select @media_globo = AVG(c) from @tblGlobo

select 
  @desc_outros = sum(vl_desconto),
  @n_desc_outros = count(vl_desconto)
from vw_movimento with (nolock) 
where vl_desconto > 0
  and cancelada = 0
  and CAST(dt_recebimento as time) > @StartTime
  and CAST(dt_recebimento as time) < @EndTime
  and CAST(dt_recebimento as date) = @Data
  and nm_desconto <> 'TV GLOBO'
-------------------------------------

----- VENDAS ------------------------
select
  @bebida = sum(i.qtd)
from vw_itemvendaunificada i with (nolock)
where i.cancelado = 0
  and i.transferido = 0  
  and i.grupo_id = 9
  and CAST(i.dt_lanc as date) = @Data
  and CAST(dt_lanc as time) > @StartTime
  and CAST(dt_lanc as time) < @EndTime
group by
  i.grupo_id

select
  @comida = sum(i.qtd)
from dbo.vw_itemvendaunificada i with (nolock)
where i.cancelado = 0
  and i.transferido = 0
  and i.grupo_id in (12)
  and CAST(i.dt_lanc as date) = @Data
  and CAST(dt_lanc as time) > @StartTime
  and CAST(dt_lanc as time) < @EndTime
group by
  i.modo_venda_descr,
  i.grupo_id,
  i.material_cod

select
  @sobremesa = sum(i.qtd)
from vw_itemvendaunificada i with (nolock)
where i.cancelado = 0
  and i.transferido = 0
  and i.grupo_id = 4
  and CAST(i.dt_lanc as date) = @Data
  and CAST(dt_lanc as time) > @StartTime
  and CAST(dt_lanc as time) < @EndTime
group by
  i.grupo_id

select
  @bomboniere = sum(i.qtd)
from vw_itemvendaunificada i with (nolock)
where i.cancelado = 0
  and i.transferido = 0
  and i.grupo_id = 11
  and CAST(i.dt_lanc as date) = @Data
  and CAST(dt_lanc as time) > @StartTime
  and CAST(dt_lanc as time) < @EndTime
group by
  i.grupo_id
------------------------------------
-- Média de Comida

declare @temp table 
(
  total_comida numeric(15,2)
)

insert @temp
	select
	  sum(i.qtd)
	from vw_itemvendaunificada i with (nolock)
	where i.cancelado = 0
	  and i.transferido = 0
	  and i.item_id = 1
	  and cast(i.dt_lanc as time) between @StartTime and @BaseTime
	  and i.dt_contabil between @BaseDate and @Data-1
	  and DATEPART(DW, dt_contabil) = DATEPART(DW, @Data)
	group by
	  i.dt_contabil

select @media_comida = avg(total_comida) from @temp

------------------------------------

declare @tblMes as table (
  total numeric(10,2)
)

insert into @tblMes
select
  total = sum(isnull(m.vl_recebido, 0))
from operacaogeral o with (nolock)
join operacaodevendageral ov with (nolock) on o.operacao_id = ov.operacao_id
join movimentocaixageral m with (nolock) on m.operacao_id = o.operacao_id
join mododevenda md with (nolock) on ov.mododevenda_id = md.id
where isnull(m.cancelado, 0) = 0
  and o.tipo = 'venda'
  and ov.encerrada = 1
  and o.cancelada = 0
  and datepart(MONTH, o.dt_contabil) = datepart(MONTH, @Data)
  and datepart(YEAR, o.dt_contabil) = datepart(YEAR, @Data)

group by o.operacao_id, ov.mododevenda_id, ov.nu_pessoas

insert into @tblMes
select
  total = sum(isnull(m.vl_recebido, 0))
from operacao o with (nolock)
join operacaodevenda ov with (nolock) on o.operacao_id = ov.operacao_id
join movimentocaixa m with (nolock) on m.operacao_id = o.operacao_id
join mododevenda md with (nolock) on ov.mododevenda_id = md.id
where isnull(m.cancelado, 0) = 0
  and o.tipo = 'venda'
  and ov.encerrada = 1
  and o.cancelada = 0
  and datepart(MONTH, o.dt_contabil) = datepart(MONTH, @Data)
  and datepart(YEAR, o.dt_contabil) = datepart(YEAR, @Data)
group by o.operacao_id, ov.mododevenda_id, ov.nu_pessoas

select
  @total_mes = sum(total)
from @tblMes

set @comida_por_pessoa = @comida / @ticket
set @media_quilo = @media_comida / @media_antes
-----------------------------------------------
SELECT @fichas_abertas = count(*) FROM [colibri8].[dbo].[ficha] where status = 'consumindo'
-----------------------------------------------

--set @ticket = @ticket - 216
--set @valor = @valor - 5700
--------------------------------------------

select  CONCAT(ISNULL(@isNovoPeriodo, ''), ISNULL(@ticket, 0), ISNULL(@isNovoPeriodo, '')) as tickets,
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
		ISNULL(@bomboniere, 0) as bomboniere,
		ISNULL(@total_mes, 0) as total_mes,
		ISNULL(@media_comida, 0) as media_comida,
		ISNULL(@comida_por_pessoa, 0) as comida_por_pessoa,
		ISNULL(@fichas_abertas, 0) as fichas_abertas,
		ISNULL(@media_quilo, 0) as media_quilo