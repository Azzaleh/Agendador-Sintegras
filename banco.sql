--
-- PostgreSQL database dump
--

\restrict NRpncU9th5ziICRTAEM5Egsh6ve0BR28sbt8edqR8JpZ1imXMsm6N82WlxUOaYx

-- Dumped from database version 18.1
-- Dumped by pg_dump version 18.1

-- Started on 2026-10-05 16:38:30

SET statement_timeout = 0;
SET lock_timeout = 0;
SET idle_in_transaction_session_timeout = 0;
SET transaction_timeout = 0;
SET client_encoding = 'UTF8';
SET standard_conforming_strings = on;
SELECT pg_catalog.set_config('search_path', '', false);
SET check_function_bodies = false;
SET xmloption = content;
SET client_min_messages = warning;
SET row_security = off;

--
-- TOC entry 2 (class 3079 OID 24585)
-- Name: pgcrypto; Type: EXTENSION; Schema: -; Owner: -
--

CREATE EXTENSION IF NOT EXISTS pgcrypto WITH SCHEMA public;


--
-- TOC entry 6079 (class 0 OID 0)
-- Dependencies: 2
-- Name: EXTENSION pgcrypto; Type: COMMENT; Schema: -; Owner: 
--

COMMENT ON EXTENSION pgcrypto IS 'cryptographic functions';


--
-- TOC entry 430 (class 1255 OID 59257)
-- Name: fn_aplicar_saldo_produto_empresa_grade_item(integer, numeric, numeric); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_aplicar_saldo_produto_empresa_grade_item(p_id_produto_empresa_grade_item integer, p_quantidade_estoque numeric, p_quantidade_prateleira numeric) RETURNS void
    LANGUAGE plpgsql
    AS $$
begin
    if coalesce(p_quantidade_estoque, 0) = 0 and coalesce(p_quantidade_prateleira, 0) = 0 then
        return;
    end if;

    perform set_config('itrade.projecao_estoque', 'S', true);

    update produto_empresa_grade_item
       set proegi_quantidade_estoque =
               coalesce(proegi_quantidade_estoque, 0) + coalesce(p_quantidade_estoque, 0),
           proegi_quantidade_prateleira =
               coalesce(proegi_quantidade_prateleira, 0) + coalesce(p_quantidade_prateleira, 0)
     where proegi_id = p_id_produto_empresa_grade_item;

    perform set_config('itrade.projecao_estoque', 'N', true);
end;
$$;


ALTER FUNCTION public.fn_aplicar_saldo_produto_empresa_grade_item(p_id_produto_empresa_grade_item integer, p_quantidade_estoque numeric, p_quantidade_prateleira numeric) OWNER TO postgres;

--
-- TOC entry 431 (class 1255 OID 50731)
-- Name: fn_atualizar_cliente_crediario(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_cliente_crediario() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    valor_total numeric(15,4);
begin
    if tg_op = 'UPDATE' then
        if new.ped_id_cliente <> old.ped_id_cliente then
	    select sum(crep_valor_total) into valor_total from crediario_parcela where crep_numero_crediario = new.ped_numero 
			and crep_status = 'P';
	    update cliente_fornecedor set cli_valor_limite_disponivel = cli_valor_limite_disponivel + valor_total  
			where clif_id = old.ped_id_cliente;  
	    update cliente_fornecedor set cli_valor_limite_disponivel = cli_valor_limite_disponivel - valor_total 
			where clif_id = new.ped_id_cliente;		
        end if;
        return new;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_cliente_crediario() OWNER TO postgres;

--
-- TOC entry 452 (class 1255 OID 25988)
-- Name: fn_atualizar_data_abertura_comanda(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_data_abertura_comanda() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    if old.com_status = 'L' and new.com_status = 'C' then
        select ped_data into new.com_data_abertura
          from pedido
         where ven_id_comanda = new.com_id
           and ped_status = 'P'
         limit 1;
    end if;

    if new.com_status = 'L' and old.com_status <> 'L' then
        new.com_data_abertura := null;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_data_abertura_comanda() OWNER TO postgres;

--
-- TOC entry 406 (class 1255 OID 59274)
-- Name: fn_atualizar_estoque_movimento_pedido(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_estoque_movimento_pedido() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_id_item integer;
begin
    if new.ped_status is distinct from old.ped_status then
        update estoque_movimento
           set estm_status = case when new.ped_status = 'X' then 'X'
                                  when new.ped_status = 'C' then 'C'
                                  else 'P' end
         where estm_numero_pedido = new.ped_numero;
    end if;

    if new.ped_tipo is distinct from old.ped_tipo then
        for v_id_item in
            select i.pedi_id
              from pedido_item i
             where i.pedi_numero_pedido = new.ped_numero
             order by i.pedi_id
        loop
            perform fn_sincronizar_estoque_movimento_pedido_item(v_id_item);
        end loop;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_estoque_movimento_pedido() OWNER TO postgres;

--
-- TOC entry 453 (class 1255 OID 59270)
-- Name: fn_atualizar_estoque_movimento_pedido_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_estoque_movimento_pedido_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    perform fn_sincronizar_estoque_movimento_pedido_item(new.pedi_id);

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_estoque_movimento_pedido_item() OWNER TO postgres;

--
-- TOC entry 372 (class 1255 OID 26000)
-- Name: fn_atualizar_quantidade_estoque_nota_fiscal(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_quantidade_estoque_nota_fiscal() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_id_movimento integer;
    v_saida boolean;
begin
    v_saida := left(new.notfi_codigo_cfop, 1) in ('5', '6');

    -- linha que nao move saldo nenhum nao existe no razao: estmi_quantidade_check
    -- recusaria a linha, e o comportamento antigo era somar zero.
    if coalesce(new.notfi_quantidade_estoque, 0) <= 0
       and (v_saida or coalesce(new.notfi_quantidade_prateleira, 0) <= 0) then
        delete from estoque_movimento_item where estmi_id_nota_fiscal_item = new.notfi_id;

        return new;
    end if;

    v_id_movimento := fn_estoque_movimento_nota_fiscal(new.notfi_id_nota_fiscal);

    insert into estoque_movimento_item (
        estmi_id_estoque_movimento, estmi_id_produto_empresa_grade_item,
        estmi_natureza, estmi_tipo,
        estmi_quantidade_estoque, estmi_quantidade_prateleira,
        estmi_id_nota_fiscal_item
    )
    values (
        v_id_movimento, new.notfi_id_produto_empresa_grade_item,
        case when v_saida then 'S' else 'E' end, 'N',
        coalesce(new.notfi_quantidade_estoque, 0),
        case when v_saida then 0 else coalesce(new.notfi_quantidade_prateleira, 0) end,
        new.notfi_id
    )
    on conflict (estmi_id_nota_fiscal_item) where estmi_id_nota_fiscal_item is not null
    do update set
        estmi_id_produto_empresa_grade_item = excluded.estmi_id_produto_empresa_grade_item,
        estmi_natureza = excluded.estmi_natureza,
        estmi_quantidade_estoque = excluded.estmi_quantidade_estoque,
        estmi_quantidade_prateleira = excluded.estmi_quantidade_prateleira;

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_quantidade_estoque_nota_fiscal() OWNER TO postgres;

--
-- TOC entry 399 (class 1255 OID 26012)
-- Name: fn_atualizar_quantidade_itens_pedido(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_quantidade_itens_pedido() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin

    if tg_op = 'INSERT' then

        update pedido
           set ped_quantidade_itens = coalesce(ped_quantidade_itens, 0) + 1
         where ped_numero = new.pedi_numero_pedido;

        return new;
    end if;

    if tg_op = 'DELETE' then

        update pedido
           set ped_quantidade_itens = greatest(coalesce(ped_quantidade_itens, 0) - 1, 0)
         where ped_numero = old.pedi_numero_pedido;

        return old;
    end if;

    return null;
end;
$$;


ALTER FUNCTION public.fn_atualizar_quantidade_itens_pedido() OWNER TO postgres;

--
-- TOC entry 390 (class 1255 OID 26022)
-- Name: fn_atualizar_status_comanda(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_status_comanda() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    if TG_OP = 'INSERT' then
        if new.ven_id_comanda is not null and new.ped_status = 'P' then
            update comanda
               set com_status = 'C'
             where com_id = new.ven_id_comanda
               and com_status = 'L';
        end if;

    elsif TG_OP = 'UPDATE' then
        -- Pedido concluído ou cancelado — libera a comanda
        if new.ped_status in ('X', 'C')
           and new.ped_status is distinct from old.ped_status
           and new.ven_id_comanda is not null
        then
            update comanda
               set com_status = 'L'
             where com_id = new.ven_id_comanda;

        -- Pedido desassociado da comanda
        elsif old.ven_id_comanda is not null
              and new.ven_id_comanda is null
        then
            update comanda
               set com_status = 'L'
             where com_id = old.ven_id_comanda;

        -- Pedido associado a uma comanda com status P
        elsif new.ven_id_comanda is not null
              and new.ped_status = 'P'
              and new.ven_id_comanda is distinct from old.ven_id_comanda
        then
            update comanda
               set com_status = 'C'
             where com_id = new.ven_id_comanda
               and com_status = 'L';
        end if;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_status_comanda() OWNER TO postgres;

--
-- TOC entry 409 (class 1255 OID 26004)
-- Name: fn_atualizar_totais_pagamento(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_totais_pagamento() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin

    -- INSERT
    if tg_op = 'INSERT' then

        update pagamento
           set pag_valor_confirmado = pag_valor_confirmado +
               case when new.pagi_status = 'C' then new.pagi_valor else 0 end,

               pag_valor_pendente = pag_valor_pendente +
               case when new.pagi_status = 'P' then new.pagi_valor else 0 end,

               pag_valor_cancelado = pag_valor_cancelado +
               case when new.pagi_status = 'X' then new.pagi_valor else 0 end

         where pag_id = new.pagi_id_pagamento;

        return null;
    end if;


    -- DELETE
    if tg_op = 'DELETE' then

        update pagamento
           set pag_valor_confirmado = pag_valor_confirmado -
               case when old.pagi_status = 'C' then old.pagi_valor else 0 end,

               pag_valor_pendente = pag_valor_pendente -
               case when old.pagi_status = 'P' then old.pagi_valor else 0 end,

               pag_valor_cancelado = pag_valor_cancelado -
               case when old.pagi_status = 'X' then old.pagi_valor else 0 end

         where pag_id = old.pagi_id_pagamento;

        return null;
    end if;


    -- UPDATE
    if tg_op = 'UPDATE' then

        update pagamento
           set pag_valor_confirmado =
               pag_valor_confirmado
               - case when old.pagi_status = 'C' then old.pagi_valor else 0 end
               + case when new.pagi_status = 'C' then new.pagi_valor else 0 end,

               pag_valor_pendente =
               pag_valor_pendente
               - case when old.pagi_status = 'P' then old.pagi_valor else 0 end
               + case when new.pagi_status = 'P' then new.pagi_valor else 0 end,

               pag_valor_cancelado =
               pag_valor_cancelado
               - case when old.pagi_status = 'X' then old.pagi_valor else 0 end
               + case when new.pagi_status = 'X' then new.pagi_valor else 0 end

         where pag_id = new.pagi_id_pagamento;

        return null;
    end if;

    return null;

end;
$$;


ALTER FUNCTION public.fn_atualizar_totais_pagamento() OWNER TO postgres;

--
-- TOC entry 418 (class 1255 OID 26024)
-- Name: fn_atualizar_valor_consumo_comanda(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_valor_consumo_comanda() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    if new.ven_id_comanda is not null then
        if new.ped_status = 'P'
           and (TG_OP = 'INSERT' or new.ped_valor_total is distinct from old.ped_valor_total)
        then
            update comanda
               set com_valor_consumo = new.ped_valor_total
             where com_id = new.ven_id_comanda;

        elsif new.ped_status in ('X', 'C')
              and (TG_OP = 'UPDATE' and new.ped_status is distinct from old.ped_status)
        then
            update comanda
               set com_valor_consumo = 0
             where com_id = new.ven_id_comanda;
        end if;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_valor_consumo_comanda() OWNER TO postgres;

--
-- TOC entry 416 (class 1255 OID 26006)
-- Name: fn_atualizar_valor_credito_cliente_por_pagamento(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_valor_credito_cliente_por_pagamento() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin

    if tg_op = 'DELETE' then
        perform fn_recalcular_credito_cliente(old.pag_id_cliente);
        return old;
    end if;

    if tg_op = 'UPDATE' and old.pag_id_cliente is distinct from new.pag_id_cliente then
        perform fn_recalcular_credito_cliente(old.pag_id_cliente);
    end if;

    perform fn_recalcular_credito_cliente(new.pag_id_cliente);

    return new;

end;
$$;


ALTER FUNCTION public.fn_atualizar_valor_credito_cliente_por_pagamento() OWNER TO postgres;

--
-- TOC entry 368 (class 1255 OID 25992)
-- Name: fn_atualizar_valor_limite_disponivel_cliente(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_valor_limite_disponivel_cliente() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_cliente_id bigint;
    v_diferenca numeric(15,2) := 0;
begin

    select p.ped_id_cliente
      into v_cliente_id
      from pedido p
     where p.ped_numero = coalesce(new.crep_numero_crediario, old.crep_numero_crediario);

    if v_cliente_id is null then
        return new;
    end if;


    if tg_op = 'INSERT' then
        
        if new.crep_status = 'P' then
            v_diferenca := new.crep_valor;
        end if;

    elsif tg_op = 'UPDATE' then
        if old.crep_status <> 'P' and new.crep_status = 'P' then
            v_diferenca := new.crep_valor;
        elsif old.crep_status = 'P' and new.crep_status in ('C','X') then
            v_diferenca := - old.crep_valor;

        elsif old.crep_status = 'P' and new.crep_status = 'P'
              and old.crep_valor <> new.crep_valor then
            v_diferenca := new.crep_valor - old.crep_valor;
        end if;

    elsif tg_op = 'DELETE' then
        
        if old.crep_status = 'P' then
            v_diferenca := - old.crep_valor;
        end if;

    end if;

    if v_diferenca <> 0 then
        update cliente_fornecedor
           set cli_valor_limite_disponivel =
               cli_valor_limite_disponivel - v_diferenca
         where clif_id = v_cliente_id;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_valor_limite_disponivel_cliente() OWNER TO postgres;

--
-- TOC entry 365 (class 1255 OID 26016)
-- Name: fn_atualizar_valores_pedido(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_atualizar_valores_pedido() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_subtotal numeric(15,4);
begin

    select coalesce(sum(i.pedi_valor_total), 0)
      into v_subtotal
      from pedido_item i
     where i.pedi_numero_pedido = new.pedi_numero_pedido
       and i.pedi_status <> 'X';

    update pedido
       set ped_valor_subtotal = v_subtotal,
           ped_valor_total    = greatest(
               v_subtotal
             + coalesce(ped_valor_acrescimo, 0)
             - coalesce(ped_valor_desconto, 0)
             + coalesce(ped_valor_frete, 0)
             + case
                   when ped_tipo = 'C' then coalesce(cred_valor_juros, 0)
                                          - coalesce(cred_valor_pagamento_previo, 0)
                   else 0
               end,
               0)
     where ped_numero = new.pedi_numero_pedido
       and ped_status <> 'X';

    return new;
end;
$$;


ALTER FUNCTION public.fn_atualizar_valores_pedido() OWNER TO postgres;

--
-- TOC entry 404 (class 1255 OID 59253)
-- Name: fn_bloquear_exclusao_estoque_movimento(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_bloquear_exclusao_estoque_movimento() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    raise exception 'Movimento de estoque não pode ser excluído; cancele o movimento';
    return old;
end;
$$;


ALTER FUNCTION public.fn_bloquear_exclusao_estoque_movimento() OWNER TO postgres;

--
-- TOC entry 383 (class 1255 OID 59255)
-- Name: fn_bloquear_exclusao_estoque_movimento_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_bloquear_exclusao_estoque_movimento_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_status varchar(1);
begin
    select m.estm_status
      into v_status
      from estoque_movimento m
     where m.estm_id = old.estmi_id_estoque_movimento;

    if v_status is distinct from 'P' then
        raise exception 'Item de movimento não pendente não pode ser excluído';
    end if;

    return old;
end;
$$;


ALTER FUNCTION public.fn_bloquear_exclusao_estoque_movimento_item() OWNER TO postgres;

--
-- TOC entry 419 (class 1255 OID 26008)
-- Name: fn_calcular_totais_pagamento(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_calcular_totais_pagamento() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin

    new.pag_valor_troco :=
      greatest(0, new.pag_valor_confirmado - new.pag_valor);

    new.pag_valor_restante :=
      greatest(0, new.pag_valor - new.pag_valor_confirmado);

    return new;

end;
$$;


ALTER FUNCTION public.fn_calcular_totais_pagamento() OWNER TO postgres;

--
-- TOC entry 395 (class 1255 OID 59266)
-- Name: fn_cancelar_estoque_movimento_nota_fiscal(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_cancelar_estoque_movimento_nota_fiscal() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    update estoque_movimento
       set estm_status = 'X'
     where estm_id_nota_fiscal = old.not_id
       and estm_status <> 'X';

    return old;
end;
$$;


ALTER FUNCTION public.fn_cancelar_estoque_movimento_nota_fiscal() OWNER TO postgres;

--
-- TOC entry 366 (class 1255 OID 59276)
-- Name: fn_cancelar_estoque_movimento_pedido(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_cancelar_estoque_movimento_pedido() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    update estoque_movimento
       set estm_status = 'X'
     where estm_numero_pedido = old.ped_numero
       and estm_status <> 'X';

    return old;
end;
$$;


ALTER FUNCTION public.fn_cancelar_estoque_movimento_pedido() OWNER TO postgres;

SET default_tablespace = '';

SET default_table_access_method = heap;

--
-- TOC entry 309 (class 1259 OID 25274)
-- Name: terminal_historico; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.terminal_historico (
    terh_id integer NOT NULL,
    terh_id_terminal integer NOT NULL,
    terh_id_funcionario_abertura integer,
    terh_id_funcionario_fechamento integer,
    terh_data date NOT NULL,
    terh_data_fechamento timestamp with time zone,
    terh_valor_abertura numeric(15,2),
    terh_valor_entrada numeric(15,2) DEFAULT 0.00,
    terh_valor_saida numeric(15,2) DEFAULT 0.00,
    terh_valor_dinheiro numeric(15,2) DEFAULT 0.00,
    terh_valor_cartao_debito numeric(15,2) DEFAULT 0.00,
    terh_valor_cartao_credito numeric(15,2) DEFAULT 0.00,
    terh_valor_pix numeric(15,2) DEFAULT 0.00,
    terh_valor_boleto_bancario numeric(15,2) DEFAULT 0.00,
    terh_valor_outros numeric(15,2) DEFAULT 0.00,
    terh_valor_venda numeric(15,2) DEFAULT 0.00,
    terh_valor_crediario numeric(15,2) DEFAULT 0.00,
    terh_valor_contas_recebidas numeric(15,2) DEFAULT 0.00,
    terh_valor_contas_pagas numeric(15,2) DEFAULT 0.00,
    terh_valor_total_movimentacao_caixa numeric(15,2) DEFAULT 0.00,
    terh_valor_total_movimentacao_banco numeric(15,2) DEFAULT 0.00,
    terh_valor_total_movimentacao numeric(15,2) DEFAULT 0.00,
    terh_valor_saldo_caixa numeric(15,2) DEFAULT 0.00,
    terh_valor_saldo_total numeric(15,2) DEFAULT 0.00,
    terh_valor_devolucao numeric(15,2) DEFAULT 0.00,
    terh_fechado boolean DEFAULT false NOT NULL,
    terh_valor_sangria numeric(15,2) DEFAULT 0.00,
    terh_valor_suprimento numeric(15,2) DEFAULT 0.00,
    terh_valor_aporte numeric(15,2) DEFAULT 0.00,
    terh_valor_recebimento numeric(15,2) DEFAULT 0.00,
    terh_valor_troco numeric(15,2) DEFAULT 0.00,
    terh_data_abertura timestamp with time zone NOT NULL,
    terh_sequencia integer DEFAULT 1 NOT NULL,
    CONSTRAINT terh_fechado_terh_data_fechamento_check CHECK ((terh_fechado = (terh_data_fechamento IS NOT NULL)))
);


ALTER TABLE public.terminal_historico OWNER TO postgres;

--
-- TOC entry 387 (class 1255 OID 25982)
-- Name: fn_consultar_terminal(integer, date); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_consultar_terminal(p_terminal_id integer, p_data date) RETURNS SETOF public.terminal_historico
    LANGUAGE plpgsql
    AS $$
DECLARE
    v_terh_id           INTEGER;
    v_linha_sintetica   terminal_historico;
BEGIN
    SELECT terh_id
      INTO v_terh_id
      FROM terminal_historico
     WHERE terh_id_terminal = p_terminal_id
       AND terh_data        = p_data
     ORDER BY terh_fechado, terh_sequencia DESC
     LIMIT 1;

    IF v_terh_id IS NOT NULL THEN
        RETURN QUERY SELECT * FROM fn_consultar_terminal_sessao(v_terh_id);
        RETURN;
    END IF;

    v_linha_sintetica.terh_id_terminal                          := p_terminal_id;
    v_linha_sintetica.terh_data                                 := p_data;
    v_linha_sintetica.terh_data_abertura                        := p_data::TIMESTAMPTZ;
    v_linha_sintetica.terh_sequencia                            := 1;
    v_linha_sintetica.terh_fechado                              := FALSE;
    v_linha_sintetica.terh_valor_abertura                       := 0;
    v_linha_sintetica.terh_valor_entrada                        := 0;
    v_linha_sintetica.terh_valor_saida                          := 0;
    v_linha_sintetica.terh_valor_total_movimentacao_caixa       := 0;
    v_linha_sintetica.terh_valor_total_movimentacao_banco       := 0;
    v_linha_sintetica.terh_valor_total_movimentacao             := 0;
    v_linha_sintetica.terh_valor_saldo_caixa                    := 0;
    v_linha_sintetica.terh_valor_saldo_total                    := 0;
    v_linha_sintetica.terh_valor_dinheiro                       := 0;
    v_linha_sintetica.terh_valor_cartao_debito                  := 0;
    v_linha_sintetica.terh_valor_cartao_credito                 := 0;
    v_linha_sintetica.terh_valor_pix                            := 0;
    v_linha_sintetica.terh_valor_boleto_bancario                := 0;
    v_linha_sintetica.terh_valor_outros                         := 0;
    v_linha_sintetica.terh_valor_venda                          := 0;
    v_linha_sintetica.terh_valor_crediario                      := 0;
    v_linha_sintetica.terh_valor_devolucao                      := 0;
    v_linha_sintetica.terh_valor_sangria                        := 0;
    v_linha_sintetica.terh_valor_recebimento                    := 0;
    v_linha_sintetica.terh_valor_troco                          := 0;
    v_linha_sintetica.terh_valor_suprimento                     := 0;
    v_linha_sintetica.terh_valor_aporte                         := 0;
    v_linha_sintetica.terh_valor_contas_recebidas               := 0;
    v_linha_sintetica.terh_valor_contas_pagas                   := 0;

    RETURN NEXT v_linha_sintetica;
END;
$$;


ALTER FUNCTION public.fn_consultar_terminal(p_terminal_id integer, p_data date) OWNER TO postgres;

--
-- TOC entry 403 (class 1255 OID 50689)
-- Name: fn_consultar_terminal_sessao(integer); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_consultar_terminal_sessao(p_terh_id integer) RETURNS SETOF public.terminal_historico
    LANGUAGE plpgsql
    AS $$
DECLARE
    v_id_terminal        INTEGER;
    v_data               DATE;
    v_valor_abertura     NUMERIC(15,4);
    v_fechado            BOOLEAN;

    v_valor_dinheiro_entrada        NUMERIC(15,4);
    v_valor_cartao_credito_entrada  NUMERIC(15,4);
    v_valor_cartao_debito_entrada   NUMERIC(15,4);
    v_valor_pix_entrada             NUMERIC(15,4);
    v_valor_boleto_entrada          NUMERIC(15,4);
    v_valor_outros_entrada          NUMERIC(15,4);
    v_valor_total_entrada           NUMERIC(15,4);

    v_valor_dinheiro_saida          NUMERIC(15,4);
    v_valor_cartao_credito_saida    NUMERIC(15,4);
    v_valor_cartao_debito_saida     NUMERIC(15,4);
    v_valor_pix_saida               NUMERIC(15,4);
    v_valor_boleto_saida            NUMERIC(15,4);
    v_valor_outros_saida            NUMERIC(15,4);
    v_valor_total_saida             NUMERIC(15,4);

    v_valor_dinheiro                NUMERIC(15,4);
    v_valor_cartao_credito          NUMERIC(15,4);
    v_valor_cartao_debito           NUMERIC(15,4);
    v_valor_pix                     NUMERIC(15,4);
    v_valor_boleto_bancario         NUMERIC(15,4);
    v_valor_outros                  NUMERIC(15,4);

    v_valor_venda                   NUMERIC(15,4);
    v_valor_crediario               NUMERIC(15,4);
    v_valor_devolucao               NUMERIC(15,4);
    v_valor_sangria                 NUMERIC(15,4);
    v_valor_recebimento             NUMERIC(15,4);
    v_valor_troco                   NUMERIC(15,4);
    v_valor_suprimento              NUMERIC(15,4);
    v_valor_aporte                  NUMERIC(15,4);
    v_valor_contas_recebidas        NUMERIC(15,4);
    v_valor_contas_pagas            NUMERIC(15,4);

    v_total_movimentacao_caixa      NUMERIC(15,4);
    v_total_movimentacao_banco      NUMERIC(15,4);
    v_total_movimentacao            NUMERIC(15,4);
    v_saldo_caixa                   NUMERIC(15,4);
    v_saldo_total                   NUMERIC(15,4);
BEGIN
    SELECT terh_id_terminal, terh_data, COALESCE(terh_valor_abertura, 0), terh_fechado
      INTO v_id_terminal, v_data, v_valor_abertura, v_fechado
      FROM terminal_historico
     WHERE terh_id = p_terh_id;

    IF NOT FOUND THEN
        RETURN;
    END IF;

    -- Fechamento congelado: devolve o gravado
    IF v_fechado THEN
        RETURN QUERY
        SELECT * FROM terminal_historico WHERE terh_id = p_terh_id;
        RETURN;
    END IF;

    -- Totais de pagamento por forma
    WITH base AS (
        SELECT
            p.pag_natureza,
            pi.pagi_forma_pagamento,
            COALESCE(pi.pagi_valor, 0) AS valor
        FROM pagamento      p
        JOIN pagamento_item pi ON pi.pagi_id_pagamento = p.pag_id
        WHERE pi.pagi_status = 'C'
          AND (
                p.pag_id_terminal_historico = p_terh_id
             OR (p.pag_id_terminal_historico IS NULL
                 AND p.pag_id_terminal = v_id_terminal
                 AND p.pag_data::DATE  = v_data)
          )
    ),
    agg AS (
        SELECT
            pag_natureza,
            pagi_forma_pagamento,
            SUM(valor) AS total_valor
        FROM base
        GROUP BY pag_natureza, pagi_forma_pagamento
    )
    SELECT
        COALESCE(SUM(CASE WHEN pag_natureza = 'E' AND pagi_forma_pagamento = 'DN'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'E' AND pagi_forma_pagamento = 'CC'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'E' AND pagi_forma_pagamento = 'CD'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'E' AND pagi_forma_pagamento = 'PX'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'E' AND pagi_forma_pagamento = 'BL'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'E' AND pagi_forma_pagamento NOT IN ('DN','CC','CD','PX','BL')   THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'E'                                                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'S' AND pagi_forma_pagamento = 'DN'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'S' AND pagi_forma_pagamento = 'CC'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'S' AND pagi_forma_pagamento = 'CD'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'S' AND pagi_forma_pagamento = 'PX'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'S' AND pagi_forma_pagamento = 'BL'                              THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'S' AND pagi_forma_pagamento NOT IN ('DN','CC','CD','PX','BL')   THEN total_valor END), 0),
        COALESCE(SUM(CASE WHEN pag_natureza = 'S'                                                              THEN total_valor END), 0)
    INTO
        v_valor_dinheiro_entrada,
        v_valor_cartao_credito_entrada,
        v_valor_cartao_debito_entrada,
        v_valor_pix_entrada,
        v_valor_boleto_entrada,
        v_valor_outros_entrada,
        v_valor_total_entrada,
        v_valor_dinheiro_saida,
        v_valor_cartao_credito_saida,
        v_valor_cartao_debito_saida,
        v_valor_pix_saida,
        v_valor_boleto_saida,
        v_valor_outros_saida,
        v_valor_total_saida
    FROM agg;

    -- Totais por tipo de operação
    SELECT
        COALESCE(SUM(pi.pagi_valor) FILTER (WHERE p.pag_tipo = 'S'), 0)::NUMERIC(15,4),
        COALESCE(SUM(pi.pagi_valor) FILTER (WHERE p.pag_tipo = 'U'), 0)::NUMERIC(15,4),
        COALESCE(SUM(pi.pagi_valor) FILTER (WHERE p.pag_tipo = 'A'), 0)::NUMERIC(15,4),
        COALESCE(SUM(pi.pagi_valor) FILTER (WHERE p.pag_tipo = 'R' AND p.pag_natureza = 'E'), 0)::NUMERIC(15,4)
      INTO v_valor_sangria, v_valor_suprimento, v_valor_aporte, v_valor_recebimento
      FROM pagamento      p
      JOIN pagamento_item pi ON pi.pagi_id_pagamento = p.pag_id
     WHERE pi.pagi_status = 'C'
       AND (
             p.pag_id_terminal_historico = p_terh_id
          OR (p.pag_id_terminal_historico IS NULL
              AND p.pag_id_terminal = v_id_terminal
              AND p.pag_data::DATE  = v_data)
       );

    -- Troco devolvido: sai sempre em espécie
    SELECT COALESCE(SUM(p.pag_valor_troco), 0)::NUMERIC(15,4)
      INTO v_valor_troco
      FROM pagamento p
     WHERE p.pag_natureza = 'E'
       AND (
             p.pag_id_terminal_historico = p_terh_id
          OR (p.pag_id_terminal_historico IS NULL
              AND p.pag_id_terminal = v_id_terminal
              AND p.pag_data::DATE  = v_data)
       );

    v_valor_dinheiro_entrada := v_valor_dinheiro_entrada - v_valor_troco;
    v_valor_total_entrada    := v_valor_total_entrada    - v_valor_troco;

    -- Venda, crediário e devolução
    SELECT
        COALESCE(SUM(CASE WHEN ped_tipo = 'V' THEN ped_valor_total END), 0),
        COALESCE(SUM(CASE WHEN ped_tipo = 'C' THEN ped_valor_total END), 0),
        COALESCE(SUM(CASE WHEN ped_tipo = 'D' THEN ped_valor_total END), 0)
      INTO v_valor_venda, v_valor_crediario, v_valor_devolucao
      FROM pedido
     WHERE ped_status IN ('S', 'C')
       AND (
             ped_id_terminal_historico = p_terh_id
          OR (ped_id_terminal_historico IS NULL
              AND ped_id_terminal = v_id_terminal
              AND ped_data::DATE  = v_data)
       );

    -- Contas recebidas e pagas
    SELECT
        COALESCE(SUM(CASE WHEN c.cnt_modalidade = 'R' THEN cp.cntpg_valor END), 0),
        COALESCE(SUM(CASE WHEN c.cnt_modalidade = 'P' THEN cp.cntpg_valor END), 0)
      INTO v_valor_contas_recebidas, v_valor_contas_pagas
      FROM conta_pagamento cp
      JOIN conta           c ON c.cnt_id = cp.cntpg_id_conta
      JOIN pagamento       p ON p.pag_id = cp.cntpg_id_pagamento
     WHERE (
             p.pag_id_terminal_historico = p_terh_id
          OR (p.pag_id_terminal_historico IS NULL
              AND p.pag_id_terminal = v_id_terminal
              AND p.pag_data::DATE  = v_data)
       );

    -- Formas de pagamento: apenas entradas
    v_valor_dinheiro        := v_valor_dinheiro_entrada;
    v_valor_cartao_credito  := v_valor_cartao_credito_entrada;
    v_valor_cartao_debito   := v_valor_cartao_debito_entrada;
    v_valor_pix             := v_valor_pix_entrada;
    v_valor_boleto_bancario := v_valor_boleto_entrada;
    v_valor_outros          := v_valor_outros_entrada;

    -- Movimentações líquidas (entrada - saída) por grupo
    v_total_movimentacao_caixa :=
        (v_valor_dinheiro_entrada       - v_valor_dinheiro_saida);

    v_total_movimentacao_banco :=
        (v_valor_cartao_credito_entrada - v_valor_cartao_credito_saida)
      + (v_valor_cartao_debito_entrada  - v_valor_cartao_debito_saida)
      + (v_valor_pix_entrada            - v_valor_pix_saida)
      + (v_valor_boleto_entrada         - v_valor_boleto_saida);

    v_total_movimentacao := v_total_movimentacao_caixa + v_total_movimentacao_banco;

    -- Saldo disponível em caixa: abertura + movimentação líquida em dinheiro
    v_saldo_caixa := v_valor_abertura + v_total_movimentacao_caixa;

    -- Saldo total da sessão
    v_saldo_total := v_valor_total_entrada - v_valor_total_saida;

    RETURN QUERY
    UPDATE terminal_historico SET
        terh_valor_entrada                        = v_valor_total_entrada,
        terh_valor_saida                          = v_valor_total_saida,
        terh_valor_total_movimentacao_caixa       = v_total_movimentacao_caixa,
        terh_valor_total_movimentacao_banco       = v_total_movimentacao_banco,
        terh_valor_total_movimentacao             = v_total_movimentacao,
        terh_valor_saldo_caixa                    = v_saldo_caixa,
        terh_valor_saldo_total                    = v_saldo_total,
        terh_valor_dinheiro                       = v_valor_dinheiro,
        terh_valor_cartao_debito                  = v_valor_cartao_debito,
        terh_valor_cartao_credito                 = v_valor_cartao_credito,
        terh_valor_pix                            = v_valor_pix,
        terh_valor_boleto_bancario                = v_valor_boleto_bancario,
        terh_valor_outros                         = v_valor_outros,
        terh_valor_venda                          = v_valor_venda,
        terh_valor_crediario                      = v_valor_crediario,
        terh_valor_devolucao                      = v_valor_devolucao,
        terh_valor_sangria                        = v_valor_sangria,
        terh_valor_recebimento                    = v_valor_recebimento,
        terh_valor_troco                          = v_valor_troco,
        terh_valor_suprimento                     = v_valor_suprimento,
        terh_valor_aporte                         = v_valor_aporte,
        terh_valor_contas_recebidas               = v_valor_contas_recebidas,
        terh_valor_contas_pagas                   = v_valor_contas_pagas
    WHERE terh_id = p_terh_id
    RETURNING *;
END;
$$;


ALTER FUNCTION public.fn_consultar_terminal_sessao(p_terh_id integer) OWNER TO postgres;

--
-- TOC entry 361 (class 1255 OID 26070)
-- Name: fn_definir_codigo_cliente_fornecedor(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_definir_codigo_cliente_fornecedor() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin

  if new.clif_codigo is null or trim(new.clif_codigo) = '' then
    new.clif_codigo := fn_proximo_codigo_cliente_fornecedor();
  end if;

  return new;

end;
$$;


ALTER FUNCTION public.fn_definir_codigo_cliente_fornecedor() OWNER TO postgres;

--
-- TOC entry 369 (class 1255 OID 26028)
-- Name: fn_definir_codigo_produto(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_definir_codigo_produto() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin

  if new.pro_codigo is null or trim(new.pro_codigo) = '' then
    new.pro_codigo := fn_proximo_codigo_produto();
  end if;

  return new;

end;
$$;


ALTER FUNCTION public.fn_definir_codigo_produto() OWNER TO postgres;

--
-- TOC entry 417 (class 1255 OID 59249)
-- Name: fn_definir_numero_estoque_movimento(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_definir_numero_estoque_movimento() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    if new.estm_tipo in ('NF', 'PD') then
        new.estm_numero := null;
        return new;
    end if;

    if new.estm_numero is null then
        select coalesce(max(m.estm_numero), 0) + 1
          into new.estm_numero
          from estoque_movimento m
         where m.estm_id_empresa = new.estm_id_empresa;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_definir_numero_estoque_movimento() OWNER TO postgres;

--
-- TOC entry 442 (class 1255 OID 59251)
-- Name: fn_definir_numero_estoque_movimento_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_definir_numero_estoque_movimento_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    if new.estmi_numero is null then
        select coalesce(max(i.estmi_numero), 0) + 1
          into new.estmi_numero
          from estoque_movimento_item i
         where i.estmi_id_estoque_movimento = new.estmi_id_estoque_movimento;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_definir_numero_estoque_movimento_item() OWNER TO postgres;

--
-- TOC entry 394 (class 1255 OID 25986)
-- Name: fn_definir_valor_limite_disponivel_cliente(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_definir_valor_limite_disponivel_cliente() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin

    if tg_op = 'INSERT' then
        new.cli_valor_limite_disponivel :=
            coalesce(new.cli_valor_limite, 0);
        return new;
    end if;

    if tg_op = 'UPDATE' then
        if new.cli_valor_limite <> old.cli_valor_limite then
            new.cli_valor_limite_disponivel :=
                old.cli_valor_limite_disponivel
                + (new.cli_valor_limite - old.cli_valor_limite);
        end if;

        return new;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_definir_valor_limite_disponivel_cliente() OWNER TO postgres;

--
-- TOC entry 449 (class 1255 OID 59262)
-- Name: fn_estoque_movimento_nota_fiscal(integer); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_estoque_movimento_nota_fiscal(p_id_nota_fiscal integer) RETURNS integer
    LANGUAGE plpgsql
    AS $$
declare
    v_id integer;
begin
    select m.estm_id
      into v_id
      from estoque_movimento m
     where m.estm_id_nota_fiscal = p_id_nota_fiscal;

    if v_id is not null then
        return v_id;
    end if;

    insert into estoque_movimento (
        estm_id_empresa, estm_tipo, estm_status,
        estm_id_funcionario, estm_id_nota_fiscal, estm_data
    )
    select nf.not_id_empresa,
           'NF',
           case when nf.not_status in ('X', 'I') then 'X'
                when nf.not_status in ('C', 'O') then 'C'
                else 'P' end,
           nf.not_id_funcionario,
           nf.not_id,
           nf.not_data_inclusao
      from nota_fiscal nf
     where nf.not_id = p_id_nota_fiscal
    returning estm_id into v_id;

    return v_id;
end;
$$;


ALTER FUNCTION public.fn_estoque_movimento_nota_fiscal(p_id_nota_fiscal integer) OWNER TO postgres;

--
-- TOC entry 396 (class 1255 OID 59268)
-- Name: fn_estoque_movimento_pedido(integer); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_estoque_movimento_pedido(p_numero_pedido integer) RETURNS integer
    LANGUAGE plpgsql
    AS $$
declare
    v_id integer;
begin
    select m.estm_id
      into v_id
      from estoque_movimento m
     where m.estm_numero_pedido = p_numero_pedido;

    if v_id is not null then
        return v_id;
    end if;

    insert into estoque_movimento (
        estm_id_empresa, estm_tipo, estm_status,
        estm_id_funcionario, estm_numero_pedido, estm_data
    )
    select t.ter_id_empresa,
           'PD',
           case when p.ped_status = 'X' then 'X'
                when p.ped_status = 'C' then 'C'
                else 'P' end,
           p.ped_id_funcionario,
           p.ped_numero,
           p.ped_data
      from pedido p
      join terminal t on t.ter_id = p.ped_id_terminal
     where p.ped_numero = p_numero_pedido
    returning estm_id into v_id;

    return v_id;
end;
$$;


ALTER FUNCTION public.fn_estoque_movimento_pedido(p_numero_pedido integer) OWNER TO postgres;

--
-- TOC entry 448 (class 1255 OID 59264)
-- Name: fn_excluir_estoque_movimento_nota_fiscal_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_excluir_estoque_movimento_nota_fiscal_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    delete from estoque_movimento_item i
     using estoque_movimento m
     where i.estmi_id_nota_fiscal_item = old.notfi_id
       and m.estm_id = i.estmi_id_estoque_movimento
       and m.estm_status <> 'X';

    return old;
end;
$$;


ALTER FUNCTION public.fn_excluir_estoque_movimento_nota_fiscal_item() OWNER TO postgres;

--
-- TOC entry 424 (class 1255 OID 59272)
-- Name: fn_excluir_estoque_movimento_pedido_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_excluir_estoque_movimento_pedido_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    delete from estoque_movimento_item i
     using estoque_movimento m
     where i.estmi_id_pedido_item = old.pedi_id
       and m.estm_id = i.estmi_id_estoque_movimento
       and m.estm_status <> 'X';

    return old;
end;
$$;


ALTER FUNCTION public.fn_excluir_estoque_movimento_pedido_item() OWNER TO postgres;

--
-- TOC entry 363 (class 1255 OID 25984)
-- Name: fn_gerar_contas_recorrentes(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_gerar_contas_recorrentes() RETURNS void
    LANGUAGE plpgsql
    AS $$
declare
    rec record;
    v_proxima_data date;
    v_novo_id int;
begin
    for rec in
        select
            cr.cntr_id_conta,
            cr.cntr_intervalo,
            c.*
        from conta_recorrencia cr
        join conta c on c.cnt_id = cr.cntr_id_conta
    loop

        v_proxima_data :=
            make_date(
                extract(year from rec.cnt_data_emissao + (rec.cntr_intervalo || ' month')::interval)::int,
                extract(month from rec.cnt_data_emissao + (rec.cntr_intervalo || ' month')::interval)::int,
                least(
                    extract(day from rec.cnt_data_emissao)::int,
                    extract(day from (
                        date_trunc('month', rec.cnt_data_emissao + (rec.cntr_intervalo || ' month')::interval)
                        + interval '1 month'
                        - interval '1 day'
                    ))::int
                )
            );

        while v_proxima_data <= current_date loop

            insert into conta (
                cnt_id_conta_tipo,
				cnt_numero_documento,
                cnt_descricao,
                cnt_data_lancamento,
                cnt_data_emissao,
                cnt_valor,
                cnt_valor_multa,
                cnt_valor_juros_mora_dia,
                cnt_id_funcionario,
                cnt_id_cliente_fornecedor,
                cnt_id_empresa,
                cnt_status,
                cnt_modalidade,
                cnt_observacao
            )
            values (
                rec.cnt_id_conta_tipo,
				'',
                rec.cnt_descricao,
                current_timestamp,
                v_proxima_data,
                rec.cnt_valor,
                rec.cnt_valor_multa,
                rec.cnt_valor_juros_mora_dia,
                rec.cnt_id_funcionario,
                rec.cnt_id_cliente_fornecedor,
                rec.cnt_id_empresa,
                rec.cnt_status,
                rec.cnt_modalidade,
                rec.cnt_observacao
            )
            returning cnt_id into v_novo_id;

            insert into conta_parcela (
                cntp_id_conta,
                cntp_numero,
                cntp_numero_parcelas,
                cntp_valor,
                cntp_valor_multa,
                cntp_valor_juros_mora,
                cntp_valor_desconto,
                cntp_valor_total,
                cntp_data_vencimento,
                cntp_data_vencimento_util,
                cntp_data_pagamento,
                cntp_status
            )
            select
                v_novo_id,
                p.cntp_numero,
                p.cntp_numero_parcelas,
                p.cntp_valor,
                p.cntp_valor_multa,
                p.cntp_valor_juros_mora,
                p.cntp_valor_desconto,
                p.cntp_valor_total,
                p.cntp_data_vencimento + (rec.cntr_intervalo || ' month')::interval,
                p.cntp_data_vencimento_util + (rec.cntr_intervalo || ' month')::interval,
                null,
                p.cntp_status
            from conta_parcela p
            where p.cntp_id_conta = rec.cnt_id;

            update conta_recorrencia
            set cntr_id_conta = v_novo_id
            where cntr_id_conta = rec.cnt_id;

            v_proxima_data :=
                make_date(
                    extract(year from v_proxima_data + (rec.cntr_intervalo || ' month')::interval)::int,
                    extract(month from v_proxima_data + (rec.cntr_intervalo || ' month')::interval)::int,
                    least(
                        extract(day from rec.cnt_data_emissao)::int,
                        extract(day from (
                            date_trunc('month', v_proxima_data + (rec.cntr_intervalo || ' month')::interval)
                            + interval '1 month'
                            - interval '1 day'
                        ))::int
                    )
                );

        end loop;

    end loop;
end;
$$;


ALTER FUNCTION public.fn_gerar_contas_recorrentes() OWNER TO postgres;

--
-- TOC entry 371 (class 1255 OID 25985)
-- Name: fn_gerar_crediario_contrato_cliente(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_gerar_crediario_contrato_cliente() RETURNS void
    LANGUAGE plpgsql
    AS $$
declare
    r record;
    v_hoje date := current_date;
    v_ultimo_dia_mes date := (date_trunc('month', current_date) + interval '1 month - 1 day')::date;
begin

    for r in
        select cc.contc_id
        from contrato_cliente cc
        join contrato c on c.cont_id = cc.contc_id_contrato
        where cc.contc_status = 'A'
          and cc.contc_data_contratacao <= v_hoje
          and (cc.contc_data_encerramento is null or cc.contc_data_encerramento >= v_hoje)
          and (
              c.cont_dia_lancamento <= extract(day from v_hoje)
              or (v_hoje = v_ultimo_dia_mes and c.cont_dia_lancamento > extract(day from v_hoje))
          )
    loop
        perform fn_gerar_crediario_contrato_cliente(r.contc_id);
    end loop;

end;
$$;


ALTER FUNCTION public.fn_gerar_crediario_contrato_cliente() OWNER TO postgres;

--
-- TOC entry 378 (class 1255 OID 34292)
-- Name: fn_gerar_crediario_contrato_cliente(integer); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_gerar_crediario_contrato_cliente(p_contc_id integer) RETURNS boolean
    LANGUAGE plpgsql
    AS $$
declare
    r record;
    v_hoje date := current_date;
    v_data_vencimento date;
    v_id_pedido integer;
    v_ultimo_dia_mes date := (date_trunc('month', current_date) + interval '1 month - 1 day')::date;
begin

    select
        cc.contc_id,
        cc.contc_id_terminal,
        cc.contc_id_cliente,
        c.cont_valor,
        c.cont_dia_lancamento,
        c.cont_dia_vencimento
    into r
    from contrato_cliente cc
    join contrato c on c.cont_id = cc.contc_id_contrato
    where cc.contc_id = p_contc_id
      and cc.contc_status in ('A', 'I')
      and cc.contc_data_contratacao <= v_hoje
      and (cc.contc_data_encerramento is null or cc.contc_data_encerramento >= v_hoje);

    if not found then
        return false;
    end if;

    if exists (
        select 1 from pedido
        where cred_id_contrato_cliente = r.contc_id
          and ped_data >= date_trunc('month', v_hoje)
          and ped_data <  date_trunc('month', v_hoje) + interval '1 month'
    ) then
        return false;
    end if;

    if r.cont_dia_vencimento < r.cont_dia_lancamento then
        v_data_vencimento := make_date(
            extract(year from v_hoje + interval '1 month')::int,
            extract(month from v_hoje + interval '1 month')::int,
            least(r.cont_dia_vencimento,
                extract(day from (date_trunc('month', v_hoje + interval '1 month') + interval '1 month - 1 day'))::int
            )
        );
    else
        v_data_vencimento := make_date(
            extract(year from v_hoje)::int,
            extract(month from v_hoje)::int,
            least(r.cont_dia_vencimento,
                extract(day from v_ultimo_dia_mes)::int
            )
        );
    end if;

    insert into pedido (
        ped_id_terminal,
        ped_id_cliente,
        ped_tipo,
        ped_status,
        ped_observacao,
        ped_data,
        ped_valor_subtotal,
        ped_valor_total,
        cred_id_contrato_cliente
    )
    values (
        r.contc_id_terminal,
        r.contc_id_cliente,
        'C',
        'P',
        '',
        v_hoje,
        r.cont_valor,
        r.cont_valor,
        r.contc_id
    ) returning ped_numero into v_id_pedido;

    insert into crediario_parcela (
        crep_numero_crediario,
        crep_numero,
        crep_numero_parcelas,
        crep_status,
        crep_valor,
        crep_valor_total,
        crep_data_vencimento
    )
    values (
        v_id_pedido,
        1,
        1,
        'P',
        r.cont_valor,
        r.cont_valor,
        v_data_vencimento
    );

    return true;

end;
$$;


ALTER FUNCTION public.fn_gerar_crediario_contrato_cliente(p_contc_id integer) OWNER TO postgres;

--
-- TOC entry 362 (class 1255 OID 25994)
-- Name: fn_inserir_parametro_empresa_apos_empresa(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_inserir_parametro_empresa_apos_empresa() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
  insert into parametro_empresa (
      pare_id_empresa,
      pare_chave_parametro
  )
  select
      new.emp_id,
      p.par_chave
  from parametro p
  on conflict do nothing;

  return new;
end;
$$;


ALTER FUNCTION public.fn_inserir_parametro_empresa_apos_empresa() OWNER TO postgres;

--
-- TOC entry 441 (class 1255 OID 26010)
-- Name: fn_inserir_parametro_empresa_apos_parametro(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_inserir_parametro_empresa_apos_parametro() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
  insert into parametro_empresa (
      pare_id_empresa,
      pare_chave_parametro
  )
  select
      e.emp_id,
      new.par_chave
  from empresa e
  on conflict do nothing;

  return new;
end;
$$;


ALTER FUNCTION public.fn_inserir_parametro_empresa_apos_parametro() OWNER TO postgres;

--
-- TOC entry 425 (class 1255 OID 25996)
-- Name: fn_inserir_produto_empresa_apos_empresa(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_inserir_produto_empresa_apos_empresa() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    insert into produto_empresa (proe_id_produto, proe_id_empresa)
    select pro_id, new.emp_id
    from produto;
    
    return null;
end;
$$;


ALTER FUNCTION public.fn_inserir_produto_empresa_apos_empresa() OWNER TO postgres;

--
-- TOC entry 408 (class 1255 OID 25998)
-- Name: fn_inserir_produto_empresa_grade_item_apos_grade_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_inserir_produto_empresa_grade_item_apos_grade_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
  insert into produto_empresa_grade_item (
      proegi_id_produto_empresa,
      proegi_id_grade_item
  )
  select
      pe.proe_id,
      new.grai_id
  from produto_empresa pe
  inner join produto p on p.pro_id = pe.proe_id_produto
  where p.pro_id_grade = new.grai_id_grade
  on conflict do nothing;

  return new;
end;
$$;


ALTER FUNCTION public.fn_inserir_produto_empresa_grade_item_apos_grade_item() OWNER TO postgres;

--
-- TOC entry 384 (class 1255 OID 26026)
-- Name: fn_inserir_produto_empresa_grade_item_apos_produto_empresa(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_inserir_produto_empresa_grade_item_apos_produto_empresa() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
  insert into produto_empresa_grade_item (
      proegi_id_produto_empresa,
      proegi_id_grade_item
  )
  select
      new.proe_id,
      gi.grai_id
  from grade_item gi
  inner join produto p on p.pro_id = new.proe_id_produto
  where gi.grai_id_grade = p.pro_id_grade
  on conflict do nothing;

  return new;
end;
$$;


ALTER FUNCTION public.fn_inserir_produto_empresa_grade_item_apos_produto_empresa() OWNER TO postgres;

--
-- TOC entry 426 (class 1255 OID 59280)
-- Name: fn_lancar_saldo_inicial_produto_empresa_grade_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_lancar_saldo_inicial_produto_empresa_grade_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_id_movimento integer;
begin
    perform set_config('itrade.projecao_estoque', 'S', true);

    update produto_empresa_grade_item
       set proegi_quantidade_estoque = 0,
           proegi_quantidade_prateleira = 0
     where proegi_id = new.proegi_id;

    perform set_config('itrade.projecao_estoque', 'N', true);

    insert into estoque_movimento (
        estm_id_empresa, estm_tipo, estm_status, estm_observacao
    )
    select pe.proe_id_empresa, 'SI', 'C', 'Saldo informado no cadastro do produto'
      from produto_empresa pe
     where pe.proe_id = new.proegi_id_produto_empresa
    returning estm_id into v_id_movimento;

    insert into estoque_movimento_item (
        estmi_id_estoque_movimento, estmi_id_produto_empresa_grade_item,
        estmi_natureza, estmi_tipo, estmi_quantidade_estoque, estmi_quantidade_prateleira
    )
    select v_id_movimento, new.proegi_id, l.natureza, 'N', l.quantidade_estoque, l.quantidade_prateleira
      from (
        select 'E' as natureza,
               greatest(coalesce(new.proegi_quantidade_estoque, 0), 0) as quantidade_estoque,
               greatest(coalesce(new.proegi_quantidade_prateleira, 0), 0) as quantidade_prateleira
        union all
        select 'S' as natureza,
               greatest(-coalesce(new.proegi_quantidade_estoque, 0), 0) as quantidade_estoque,
               greatest(-coalesce(new.proegi_quantidade_prateleira, 0), 0) as quantidade_prateleira
      ) l
     where l.quantidade_estoque > 0
        or l.quantidade_prateleira > 0;

    return new;
end;
$$;


ALTER FUNCTION public.fn_lancar_saldo_inicial_produto_empresa_grade_item() OWNER TO postgres;

--
-- TOC entry 375 (class 1255 OID 34267)
-- Name: fn_liberar_venda_pendente_terminal(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_liberar_venda_pendente_terminal() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    if new.ped_status in ('C', 'X')
       and new.ped_status is distinct from old.ped_status
    then
        update terminal
           set ter_numero_venda_pendente = null
         where ter_numero_venda_pendente = new.ped_numero;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_liberar_venda_pendente_terminal() OWNER TO postgres;

--
-- TOC entry 420 (class 1255 OID 50733)
-- Name: fn_pagamento_gera_credito(character varying, character varying); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_pagamento_gera_credito(p_tipo character varying, p_natureza character varying) RETURNS boolean
    LANGUAGE sql IMMUTABLE
    AS $$
    select coalesce(p_natureza, 'E') = 'E'
       and coalesce(p_tipo, 'V') in ('V', 'R', 'A');
$$;


ALTER FUNCTION public.fn_pagamento_gera_credito(p_tipo character varying, p_natureza character varying) OWNER TO postgres;

--
-- TOC entry 382 (class 1255 OID 59260)
-- Name: fn_projetar_estoque_movimento(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_projetar_estoque_movimento() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_fator integer;
begin
    if new.estm_status is not distinct from old.estm_status then
        return new;
    end if;

    if old.estm_status = 'X' and new.estm_status <> 'X' then
        v_fator := 1;
    elsif old.estm_status <> 'X' and new.estm_status = 'X' then
        v_fator := -1;
    else
        return new;
    end if;

    perform set_config('itrade.projecao_estoque', 'S', true);

    update produto_empresa_grade_item pgi
       set proegi_quantidade_estoque =
               coalesce(pgi.proegi_quantidade_estoque, 0) + v_fator * s.quantidade_estoque,
           proegi_quantidade_prateleira =
               coalesce(pgi.proegi_quantidade_prateleira, 0) + v_fator * s.quantidade_prateleira
      from (
        select i.estmi_id_produto_empresa_grade_item as id_produto_empresa_grade_item,
               sum(case when i.estmi_natureza = 'E' then 1 else -1 end
                   * coalesce(i.estmi_quantidade_estoque, 0)) as quantidade_estoque,
               sum(case when i.estmi_natureza = 'E' then 1 else -1 end
                   * coalesce(i.estmi_quantidade_prateleira, 0)) as quantidade_prateleira
          from estoque_movimento_item i
         where i.estmi_id_estoque_movimento = new.estm_id
         group by i.estmi_id_produto_empresa_grade_item
      ) s
     where pgi.proegi_id = s.id_produto_empresa_grade_item;

    perform set_config('itrade.projecao_estoque', 'N', true);

    return new;
end;
$$;


ALTER FUNCTION public.fn_projetar_estoque_movimento() OWNER TO postgres;

--
-- TOC entry 370 (class 1255 OID 59258)
-- Name: fn_projetar_estoque_movimento_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_projetar_estoque_movimento_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_status varchar(1);
    v_fator integer;
begin
    if tg_op in ('UPDATE', 'DELETE') then
        select m.estm_status
          into v_status
          from estoque_movimento m
         where m.estm_id = old.estmi_id_estoque_movimento;

        if v_status is distinct from 'X' then
            v_fator := case when old.estmi_natureza = 'E' then -1 else 1 end;

            perform fn_aplicar_saldo_produto_empresa_grade_item(
                old.estmi_id_produto_empresa_grade_item,
                v_fator * coalesce(old.estmi_quantidade_estoque, 0),
                v_fator * coalesce(old.estmi_quantidade_prateleira, 0));
        end if;
    end if;

    if tg_op = 'DELETE' then
        return old;
    end if;

    select m.estm_status
      into v_status
      from estoque_movimento m
     where m.estm_id = new.estmi_id_estoque_movimento;

    if v_status is distinct from 'X' then
        v_fator := case when new.estmi_natureza = 'E' then 1 else -1 end;

        perform fn_aplicar_saldo_produto_empresa_grade_item(
            new.estmi_id_produto_empresa_grade_item,
            v_fator * coalesce(new.estmi_quantidade_estoque, 0),
            v_fator * coalesce(new.estmi_quantidade_prateleira, 0));
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_projetar_estoque_movimento_item() OWNER TO postgres;

--
-- TOC entry 433 (class 1255 OID 59278)
-- Name: fn_proteger_saldo_produto_empresa_grade_item(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_proteger_saldo_produto_empresa_grade_item() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    if current_setting('itrade.projecao_estoque', true) is distinct from 'S' then
        new.proegi_quantidade_estoque := old.proegi_quantidade_estoque;
        new.proegi_quantidade_prateleira := old.proegi_quantidade_prateleira;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_proteger_saldo_produto_empresa_grade_item() OWNER TO postgres;

--
-- TOC entry 393 (class 1255 OID 34259)
-- Name: fn_proximo_codigo_cliente_fornecedor(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_proximo_codigo_cliente_fornecedor() RETURNS character varying
    LANGUAGE plpgsql
    AS $$
declare
  v_codigo bigint;
begin

  loop
    v_codigo := nextval('cliente_fornecedor_clif_codigo_seq');

    exit when not exists (
      select 1 from cliente_fornecedor
      where clif_codigo = v_codigo::varchar
    );

  end loop;

  return v_codigo::varchar;

end;
$$;


ALTER FUNCTION public.fn_proximo_codigo_cliente_fornecedor() OWNER TO postgres;

--
-- TOC entry 447 (class 1255 OID 34258)
-- Name: fn_proximo_codigo_produto(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_proximo_codigo_produto() RETURNS character varying
    LANGUAGE plpgsql
    AS $$
declare
  v_codigo bigint;
begin

  loop
    v_codigo := nextval('produto_pro_codigo_seq');

    exit when not exists (
      select 1 from produto
      where pro_codigo = v_codigo::varchar
    );

  end loop;

  return v_codigo::varchar;

end;
$$;


ALTER FUNCTION public.fn_proximo_codigo_produto() OWNER TO postgres;

--
-- TOC entry 402 (class 1255 OID 50719)
-- Name: fn_recalcular_credito_cliente(integer); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_recalcular_credito_cliente(p_id_cliente integer) RETURNS void
    LANGUAGE plpgsql
    AS $$
declare
    v_credito numeric(15,4);
begin

    if p_id_cliente is null then
        return;
    end if;

    select coalesce(sum(greatest(0, pag_valor_credito)), 0)
      into v_credito
      from pagamento
     where pag_id_cliente = p_id_cliente;

    update cliente_fornecedor
       set cli_valor_credito = v_credito
     where clif_id = p_id_cliente
       and cli_valor_credito is distinct from v_credito;

end;
$$;


ALTER FUNCTION public.fn_recalcular_credito_cliente(p_id_cliente integer) OWNER TO postgres;

--
-- TOC entry 374 (class 1255 OID 50717)
-- Name: fn_recalcular_credito_pagamento(integer); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_recalcular_credito_pagamento(p_id_pagamento integer) RETURNS void
    LANGUAGE plpgsql
    AS $$
declare
    v_pagamento record;
    v_liquidado numeric(15,4);
    v_credito   numeric(15,4);
begin

    if p_id_pagamento is null then
        return;
    end if;

    select pag_valor_confirmado, pag_valor_troco, pag_valor_credito, pag_tipo, pag_natureza
      into v_pagamento
      from pagamento
     where pag_id = p_id_pagamento
       for update;

    if not found then
        return;
    end if;

    if not fn_pagamento_gera_credito(v_pagamento.pag_tipo, v_pagamento.pag_natureza) then

        v_credito := 0;

    else

        select coalesce((select sum(pedpg_valor)
                           from pedido_pagamento
                          where pedpg_id_pagamento = p_id_pagamento
                            and pedpg_status = 'C'), 0)
             + coalesce((select sum(cntpg_valor)
                           from conta_pagamento
                          where cntpg_id_pagamento = p_id_pagamento
                            and cntpg_status = 'C'), 0)
          into v_liquidado;

        v_credito := v_pagamento.pag_valor_confirmado
                   - v_pagamento.pag_valor_troco
                   - v_liquidado;

    end if;

    if v_credito is distinct from v_pagamento.pag_valor_credito then
        update pagamento
           set pag_valor_credito = v_credito
         where pag_id = p_id_pagamento;
    end if;

end;
$$;


ALTER FUNCTION public.fn_recalcular_credito_pagamento(p_id_pagamento integer) OWNER TO postgres;

--
-- TOC entry 412 (class 1255 OID 50718)
-- Name: fn_recalcular_credito_pagamento_trigger(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_recalcular_credito_pagamento_trigger() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_id_novo   integer;
    v_id_antigo integer;
begin

    if tg_table_name = 'pagamento_item' then
        if tg_op <> 'DELETE' then v_id_novo   := new.pagi_id_pagamento; end if;
        if tg_op <> 'INSERT' then v_id_antigo := old.pagi_id_pagamento; end if;

    elsif tg_table_name = 'pedido_pagamento' then
        if tg_op <> 'DELETE' then v_id_novo   := new.pedpg_id_pagamento; end if;
        if tg_op <> 'INSERT' then v_id_antigo := old.pedpg_id_pagamento; end if;

    elsif tg_table_name = 'conta_pagamento' then
        if tg_op <> 'DELETE' then v_id_novo   := new.cntpg_id_pagamento; end if;
        if tg_op <> 'INSERT' then v_id_antigo := old.cntpg_id_pagamento; end if;

    else
        return null;
    end if;

    if v_id_antigo is not null and v_id_antigo is distinct from v_id_novo then
        perform fn_recalcular_credito_pagamento(v_id_antigo);
    end if;

    perform fn_recalcular_credito_pagamento(v_id_novo);

    return null;

end;
$$;


ALTER FUNCTION public.fn_recalcular_credito_pagamento_trigger() OWNER TO postgres;

--
-- TOC entry 385 (class 1255 OID 59269)
-- Name: fn_sincronizar_estoque_movimento_pedido_item(integer); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_sincronizar_estoque_movimento_pedido_item(p_id_pedido_item integer) RETURNS void
    LANGUAGE plpgsql
    AS $$
declare
    v_item pedido_item%rowtype;
    v_tipo varchar(1);
    v_id_movimento integer;
begin
    select *
      into v_item
      from pedido_item i
     where i.pedi_id = p_id_pedido_item;

    select p.ped_tipo
      into v_tipo
      from pedido p
     where p.ped_numero = v_item.pedi_numero_pedido;

    if v_tipo not in ('V', 'C', 'D')
       or v_item.pedi_status = 'X'
       or coalesce(v_item.pedi_quantidade, 0) <= 0 then
        delete from estoque_movimento_item i
         using estoque_movimento m
         where i.estmi_id_pedido_item = p_id_pedido_item
           and m.estm_id = i.estmi_id_estoque_movimento
           and m.estm_status <> 'X';

        return;
    end if;

    v_id_movimento := fn_estoque_movimento_pedido(v_item.pedi_numero_pedido);

    delete from estoque_movimento_item i
     using estoque_movimento m
     where i.estmi_id_pedido_item = p_id_pedido_item
       and i.estmi_tipo = 'N'
       and i.estmi_id_produto_empresa_grade_item <> v_item.pedi_id_produto_empresa_grade_item
       and m.estm_id = i.estmi_id_estoque_movimento
       and m.estm_status <> 'X';

    insert into estoque_movimento_item (
        estmi_id_estoque_movimento, estmi_id_produto_empresa_grade_item,
        estmi_natureza, estmi_tipo,
        estmi_quantidade_estoque, estmi_quantidade_prateleira,
        estmi_id_pedido_item
    )
    values (
        v_id_movimento, v_item.pedi_id_produto_empresa_grade_item,
        case when v_tipo = 'D' then 'E' else 'S' end, 'N',
        0, v_item.pedi_quantidade,
        v_item.pedi_id
    )
    on conflict (estmi_id_pedido_item, estmi_id_produto_empresa_grade_item, estmi_tipo)
       where estmi_id_pedido_item is not null
    do update set
        estmi_natureza = excluded.estmi_natureza,
        estmi_quantidade_prateleira = excluded.estmi_quantidade_prateleira;
end;
$$;


ALTER FUNCTION public.fn_sincronizar_estoque_movimento_pedido_item(p_id_pedido_item integer) OWNER TO postgres;

--
-- TOC entry 364 (class 1255 OID 26018)
-- Name: fn_validar_comanda_bloqueada(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_validar_comanda_bloqueada() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_status char(1);
begin
    select com_status into v_status
      from comanda
      join pedido on com_id = ven_id_comanda
     where pedido.ped_numero = new.pedi_numero_pedido;

    if v_status = 'B' then
        raise exception 'Comanda bloqueada. Não é possível adicionar itens.';
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_validar_comanda_bloqueada() OWNER TO postgres;

--
-- TOC entry 367 (class 1255 OID 50723)
-- Name: fn_validar_credito_pagamento(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_validar_credito_pagamento() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
declare
    v_pagamento record;
begin

    select pag_id, pag_valor_credito, pag_valor_pendente, pag_id_cliente, pag_tipo, pag_natureza
      into v_pagamento
      from pagamento
     where pag_id = new.pag_id;

    if not found then
        return null;
    end if;

    if not fn_pagamento_gera_credito(v_pagamento.pag_tipo, v_pagamento.pag_natureza) then
        return null;
    end if;

    if v_pagamento.pag_valor_credito + v_pagamento.pag_valor_pendente < 0 then
        raise exception
            'Liquidação excede o valor recebido no pagamento % (crédito %, pendente %)',
            v_pagamento.pag_id,
            v_pagamento.pag_valor_credito,
            v_pagamento.pag_valor_pendente;
    end if;

    if v_pagamento.pag_valor_credito > 0 and v_pagamento.pag_id_cliente is null then
        raise exception
            'Pagamento % gerou crédito de % sem cliente vinculado',
            v_pagamento.pag_id,
            v_pagamento.pag_valor_credito;
    end if;

    return null;

end;
$$;


ALTER FUNCTION public.fn_validar_credito_pagamento() OWNER TO postgres;

--
-- TOC entry 410 (class 1255 OID 26002)
-- Name: fn_verificar_cancelamento_nota_fiscal(); Type: FUNCTION; Schema: public; Owner: postgres
--

CREATE FUNCTION public.fn_verificar_cancelamento_nota_fiscal() RETURNS trigger
    LANGUAGE plpgsql
    AS $$
begin
    if new.not_status is not distinct from old.not_status then
        return new;
    end if;

    update estoque_movimento
       set estm_status = case when new.not_status in ('X', 'I') then 'X'
                              when new.not_status in ('C', 'O') then 'C'
                              else 'P' end
     where estm_id_nota_fiscal = new.not_id;

    if new.not_status = 'X' and old.not_status <> 'X' then
        update pedido set ped_id_nota_fiscal = null where ped_id_nota_fiscal = new.not_id;
    end if;

    return new;
end;
$$;


ALTER FUNCTION public.fn_verificar_cancelamento_nota_fiscal() OWNER TO postgres;

--
-- TOC entry 227 (class 1259 OID 24656)
-- Name: acesso; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.acesso (
    ace_id_funcionario integer NOT NULL,
    ace_id_empresa integer NOT NULL
);


ALTER TABLE public.acesso OWNER TO postgres;

--
-- TOC entry 333 (class 1259 OID 25471)
-- Name: ajuda; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.ajuda (
    ajuda_codigo character varying(60) NOT NULL,
    ajuda_descricao character varying(1024) NOT NULL
);


ALTER TABLE public.ajuda OWNER TO postgres;

--
-- TOC entry 316 (class 1259 OID 25337)
-- Name: aliquota; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.aliquota (
    aliq_uf_origem character varying(2) NOT NULL,
    aliq_uf_destino character varying(2) NOT NULL,
    aliq_valor numeric(15,2)
);


ALTER TABLE public.aliquota OWNER TO postgres;

--
-- TOC entry 321 (class 1259 OID 25368)
-- Name: boleto; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.boleto (
    bol_id integer NOT NULL,
    bol_codigo_banco character varying(3) NOT NULL,
    bol_numero_cliente character varying(20) NOT NULL,
    bol_modalidade character varying(1) NOT NULL,
    bol_numero_conta_corrente character varying(10) NOT NULL,
    bol_especie character varying(3) NOT NULL,
    bol_data_emissao timestamp with time zone NOT NULL,
    bol_nosso_numero character varying(16) NOT NULL,
    bol_seu_numero character varying(16) NOT NULL,
    bol_codigo_barras character varying(44) NOT NULL,
    bol_linha_digitavel character varying(48) NOT NULL,
    bol_identificacao_emissao character varying(1),
    bol_identificacao_distribuicao character varying(1),
    bol_valor numeric(15,2) NOT NULL,
    bol_data_vencimento timestamp with time zone NOT NULL,
    bol_data_limite_pagamento timestamp with time zone,
    bol_valor_abatimento numeric(15,2) NOT NULL,
    bol_tipo_desconto character varying(1) NOT NULL,
    bol_data_primeiro_desconto timestamp with time zone,
    bol_valor_primeiro_desconto numeric(15,2),
    bol_data_segundo_desconto timestamp with time zone,
    bol_valor_segundo_desconto numeric(15,2),
    bol_data_terceiro_desconto timestamp with time zone,
    bol_valor_terceiro_desconto numeric(15,2),
    bol_tipo_multa character varying(1) NOT NULL,
    bol_data_multa timestamp with time zone,
    bol_valor_multa numeric(15,2),
    bol_tipo_juros_mora character varying(1) NOT NULL,
    bol_data_juros_mora timestamp with time zone,
    bol_valor_juros_mora numeric(15,2),
    bol_codigo_negativacao character varying(1) NOT NULL,
    bol_numero_dias_negativacao integer,
    bol_codigo_protesto character varying(1) NOT NULL,
    bol_numero_dias_protesto integer,
    bol_pagador_cpf_cnpj character varying(14),
    bol_pagador_nome character varying(100),
    bol_pagador_bairro character varying(30),
    bol_pagador_cidade character varying(40),
    bol_pagador_logradouro character varying(40),
    bol_pagador_uf character varying(2),
    bol_pagador_cep character varying(8),
    bol_pagador_email character varying(255),
    bol_beneficiario_cpf_cnpj character varying(14),
    bol_beneficiario_nome character varying(40),
    bol_aceite boolean NOT NULL,
    bol_numero_contrato integer,
    bol_situacao integer NOT NULL,
    bol_id_funcionario_emissao integer NOT NULL,
    bol_id_funcionario_baixa integer,
    bol_id_crediario_parcela integer NOT NULL
);


ALTER TABLE public.boleto OWNER TO postgres;

--
-- TOC entry 320 (class 1259 OID 25367)
-- Name: boleto_bol_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.boleto ALTER COLUMN bol_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.boleto_bol_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 337 (class 1259 OID 25980)
-- Name: boleto_bol_nosso_numero_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.boleto_bol_nosso_numero_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE public.boleto_bol_nosso_numero_seq OWNER TO postgres;

--
-- TOC entry 323 (class 1259 OID 25401)
-- Name: boleto_instrucao; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.boleto_instrucao (
    boli_id integer NOT NULL,
    boli_id_boleto integer NOT NULL,
    boli_instrucao character varying(255) NOT NULL
);


ALTER TABLE public.boleto_instrucao OWNER TO postgres;

--
-- TOC entry 322 (class 1259 OID 25400)
-- Name: boleto_instrucao_boli_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.boleto_instrucao ALTER COLUMN boli_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.boleto_instrucao_boli_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 280 (class 1259 OID 25024)
-- Name: cest; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.cest (
    cest_codigo character varying(7) NOT NULL,
    cest_ncm character varying(8) NOT NULL,
    cest_valor_mva numeric(10,2) DEFAULT 0,
    cest_descricao character varying(255)
);


ALTER TABLE public.cest OWNER TO postgres;

--
-- TOC entry 281 (class 1259 OID 25032)
-- Name: cfop; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.cfop (
    cfop_codigo character varying(4) NOT NULL,
    cfop_descricao character varying(512),
    cfop_tipo character varying(1)
);


ALTER TABLE public.cfop OWNER TO postgres;

--
-- TOC entry 303 (class 1259 OID 25244)
-- Name: classificacao_tributaria; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.classificacao_tributaria (
    ctrib_codigo character varying(3) NOT NULL,
    ctrib_descricao character varying(100) NOT NULL
);


ALTER TABLE public.classificacao_tributaria OWNER TO postgres;

--
-- TOC entry 305 (class 1259 OID 25252)
-- Name: classificacao_tributaria_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.classificacao_tributaria_item (
    ctribi_id integer NOT NULL,
    ctribi_codigo character varying(3),
    ctribi_codigo_classificacao_tributaria character varying(3) CONSTRAINT classificacao_tributaria_it_ctribi_codigo_classificaca_not_null NOT NULL,
    ctribi_descricao character varying(1024) NOT NULL
);


ALTER TABLE public.classificacao_tributaria_item OWNER TO postgres;

--
-- TOC entry 304 (class 1259 OID 25251)
-- Name: classificacao_tributaria_item_ctribi_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.classificacao_tributaria_item ALTER COLUMN ctribi_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.classificacao_tributaria_item_ctribi_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 248 (class 1259 OID 24774)
-- Name: cliente_endereco; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.cliente_endereco (
    clie_id integer NOT NULL,
    clie_id_cliente integer NOT NULL,
    clie_descricao character varying(100) NOT NULL,
    clie_principal boolean DEFAULT false,
    clie_cep character varying(8),
    clie_logradouro character varying(100),
    clie_bairro character varying(80),
    clie_cidade character varying(80),
    clie_uf character varying(2),
    clie_numero character varying(15),
    clie_complemento character varying(20),
    clie_referencia character varying(255),
    clie_codigo_municipio integer
);


ALTER TABLE public.cliente_endereco OWNER TO postgres;

--
-- TOC entry 247 (class 1259 OID 24773)
-- Name: cliente_endereco_clie_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.cliente_endereco ALTER COLUMN clie_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.cliente_endereco_clie_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 239 (class 1259 OID 24718)
-- Name: cliente_fornecedor; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.cliente_fornecedor (
    clif_id integer NOT NULL,
    clif_cpf_cnpj character varying(14),
    clif_nome character varying(100) NOT NULL,
    clif_apelido character varying(40),
    clif_inscricao_estadual character varying(18),
    clif_data_nascimento date,
    clif_data_cadastro timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    clif_telefone_principal character varying(15),
    clif_telefone_secundario character varying(15),
    clif_telefone_terciario character varying(15),
    clif_contato character varying(50),
    clif_email character varying(255),
    clif_nacionalidade character varying(20),
    clif_naturalidade character varying(80),
    clif_suframa character varying(15),
    clif_observacao character varying(255),
    clif_codigo_pais character varying(4) NOT NULL,
    clif_contribuinte boolean NOT NULL,
    clif_id_funcionario_cadastro integer NOT NULL,
    clif_tipo_cadastro character varying(1) NOT NULL,
    cli_rg character varying(15),
    cli_pai character varying(100),
    cli_mae character varying(100),
    cli_estado_civil character varying(15),
    cli_situacao character varying(2),
    cli_quantidade_dias_prazo integer DEFAULT 0,
    cli_valor_limite numeric(10,2) DEFAULT 0.00,
    cli_valor_limite_disponivel numeric(10,2) DEFAULT 0.00,
    cli_valor_credito numeric(15,2) DEFAULT 0.00,
    cli_id_cliente_grupo integer,
    cli_id_cliente_rota integer,
    for_cep character varying(8),
    for_logradouro character varying(100),
    for_numero character varying(15),
    for_complemento character varying(20),
    for_bairro character varying(80),
    for_cidade character varying(80),
    for_codigo_municipio integer,
    for_uf character varying(2),
    for_banco_titular character varying(60),
    for_banco_operacao character varying(10),
    for_banco_nome character varying(60),
    for_banco_cpf_cnpj character varying(14),
    for_banco_numero character varying(10),
    for_banco_conta character varying(15),
    for_banco_agencia character varying(10),
    clif_codigo character varying(20) NOT NULL
);


ALTER TABLE public.cliente_fornecedor OWNER TO postgres;

--
-- TOC entry 341 (class 1259 OID 26066)
-- Name: cliente_fornecedor_clif_codigo_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.cliente_fornecedor_clif_codigo_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE public.cliente_fornecedor_clif_codigo_seq OWNER TO postgres;

--
-- TOC entry 238 (class 1259 OID 24717)
-- Name: cliente_fornecedor_clif_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.cliente_fornecedor ALTER COLUMN clif_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.cliente_fornecedor_clif_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 246 (class 1259 OID 24764)
-- Name: cliente_grupo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.cliente_grupo (
    clig_id integer NOT NULL,
    clig_descricao character varying(100) NOT NULL
);


ALTER TABLE public.cliente_grupo OWNER TO postgres;

--
-- TOC entry 245 (class 1259 OID 24763)
-- Name: cliente_grupo_clig_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.cliente_grupo ALTER COLUMN clig_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.cliente_grupo_clig_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 244 (class 1259 OID 24756)
-- Name: cliente_referencia_comercial; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.cliente_referencia_comercial (
    clirc_id integer NOT NULL,
    clirc_id_cliente integer NOT NULL,
    clirc_empresa character varying(100),
    clirc_telefone_principal character varying(15),
    clirc_telefone_secundario character varying(15),
    clirc_contato character varying(50)
);


ALTER TABLE public.cliente_referencia_comercial OWNER TO postgres;

--
-- TOC entry 243 (class 1259 OID 24755)
-- Name: cliente_referencia_comercial_clirc_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.cliente_referencia_comercial ALTER COLUMN clirc_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.cliente_referencia_comercial_clirc_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 242 (class 1259 OID 24746)
-- Name: cliente_rota; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.cliente_rota (
    clir_id integer NOT NULL,
    clir_descricao character varying(100) NOT NULL
);


ALTER TABLE public.cliente_rota OWNER TO postgres;

--
-- TOC entry 241 (class 1259 OID 24745)
-- Name: cliente_rota_clir_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.cliente_rota ALTER COLUMN clir_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.cliente_rota_clir_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 334 (class 1259 OID 25480)
-- Name: comanda; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.comanda (
    com_numero integer NOT NULL,
    com_status character varying(1) NOT NULL,
    com_id_cliente integer,
    com_valor_consumo numeric(15,4) DEFAULT 0,
    com_data_abertura timestamp with time zone,
    com_id_empresa integer NOT NULL,
    com_id integer NOT NULL
);


ALTER TABLE public.comanda OWNER TO postgres;

--
-- TOC entry 342 (class 1259 OID 34270)
-- Name: comanda_com_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.comanda ALTER COLUMN com_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.comanda_com_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 325 (class 1259 OID 25410)
-- Name: conta; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.conta (
    cnt_id integer NOT NULL,
    cnt_id_conta_tipo integer,
    cnt_numero_documento character varying(20) NOT NULL,
    cnt_descricao character varying(100) NOT NULL,
    cnt_data_lancamento timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    cnt_data_emissao timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    cnt_valor numeric(15,2) NOT NULL,
    cnt_valor_multa numeric(15,2) DEFAULT 0.00,
    cnt_valor_juros_mora_dia numeric(15,4) DEFAULT 0.00,
    cnt_id_funcionario integer NOT NULL,
    cnt_id_cliente_fornecedor integer NOT NULL,
    cnt_id_empresa integer NOT NULL,
    cnt_status character varying(1) NOT NULL,
    cnt_modalidade character varying(1) NOT NULL,
    cnt_observacao character varying(255)
);


ALTER TABLE public.conta OWNER TO postgres;

--
-- TOC entry 319 (class 1259 OID 25351)
-- Name: conta_bancaria; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.conta_bancaria (
    conb_id integer NOT NULL,
    conb_banco character varying(3) NOT NULL,
    conb_agencia character varying(5) NOT NULL,
    conb_numero_cliente character varying(20),
    conb_numero character varying(10) NOT NULL,
    conb_digito_verificador character varying(1) NOT NULL,
    conb_tipo integer NOT NULL,
    conb_descricao character varying(40) NOT NULL,
    conb_observacao character varying(255),
    conb_codigo_integracao character varying(255)
);


ALTER TABLE public.conta_bancaria OWNER TO postgres;

--
-- TOC entry 318 (class 1259 OID 25350)
-- Name: conta_bancaria_conb_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.conta_bancaria ALTER COLUMN conb_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.conta_bancaria_conb_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 340 (class 1259 OID 26031)
-- Name: conta_bancaria_retorno; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.conta_bancaria_retorno (
    conbr_id integer NOT NULL,
    conbr_id_conta_bancaria integer NOT NULL,
    conbr_tipo_movimento integer NOT NULL,
    conbr_data_inicial date NOT NULL,
    conbr_data_final date NOT NULL,
    conbr_codigo_solicitacao integer NOT NULL,
    conbr_situacao integer NOT NULL,
    conbr_data_solicitacao timestamp with time zone NOT NULL,
    conbr_data_processamento timestamp with time zone,
    conbr_id_funcionario integer
);


ALTER TABLE public.conta_bancaria_retorno OWNER TO postgres;

--
-- TOC entry 339 (class 1259 OID 26030)
-- Name: conta_bancaria_retorno_conbr_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.conta_bancaria_retorno ALTER COLUMN conbr_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.conta_bancaria_retorno_conbr_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 324 (class 1259 OID 25409)
-- Name: conta_cnt_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.conta ALTER COLUMN cnt_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.conta_cnt_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 338 (class 1259 OID 25981)
-- Name: conta_cnt_numero_documento_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.conta_cnt_numero_documento_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE public.conta_cnt_numero_documento_seq OWNER TO postgres;

--
-- TOC entry 329 (class 1259 OID 25443)
-- Name: conta_pagamento; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.conta_pagamento (
    cntpg_id integer NOT NULL,
    cntpg_id_conta integer NOT NULL,
    cntpg_id_pagamento integer NOT NULL,
    cntpg_valor numeric(15,2) NOT NULL,
    cntpg_status character varying(1) NOT NULL,
    cntpg_data timestamp with time zone DEFAULT CURRENT_TIMESTAMP
);


ALTER TABLE public.conta_pagamento OWNER TO postgres;

--
-- TOC entry 328 (class 1259 OID 25442)
-- Name: conta_pagamento_cntpg_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.conta_pagamento ALTER COLUMN cntpg_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.conta_pagamento_cntpg_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 327 (class 1259 OID 25429)
-- Name: conta_parcela; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.conta_parcela (
    cntp_id integer NOT NULL,
    cntp_id_conta integer,
    cntp_numero integer NOT NULL,
    cntp_numero_parcelas integer NOT NULL,
    cntp_valor numeric(15,2) DEFAULT 0.00,
    cntp_valor_multa numeric(15,2) DEFAULT 0.00,
    cntp_valor_juros_mora numeric(15,2) DEFAULT 0.00,
    cntp_valor_desconto numeric(15,2) DEFAULT 0.00,
    cntp_valor_total numeric(15,2) NOT NULL,
    cntp_data_vencimento date,
    cntp_data_vencimento_util date,
    cntp_data_pagamento timestamp with time zone,
    cntp_status character varying(1)
);


ALTER TABLE public.conta_parcela OWNER TO postgres;

--
-- TOC entry 326 (class 1259 OID 25428)
-- Name: conta_parcela_cntp_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.conta_parcela ALTER COLUMN cntp_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.conta_parcela_cntp_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 332 (class 1259 OID 25464)
-- Name: conta_recorrencia; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.conta_recorrencia (
    cntr_id_conta integer NOT NULL,
    cntr_intervalo integer NOT NULL
);


ALTER TABLE public.conta_recorrencia OWNER TO postgres;

--
-- TOC entry 331 (class 1259 OID 25457)
-- Name: conta_tipo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.conta_tipo (
    cntt_id integer NOT NULL,
    cntt_descricao character varying(100) NOT NULL
);


ALTER TABLE public.conta_tipo OWNER TO postgres;

--
-- TOC entry 330 (class 1259 OID 25456)
-- Name: conta_tipo_cntt_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.conta_tipo ALTER COLUMN cntt_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.conta_tipo_cntt_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 311 (class 1259 OID 25304)
-- Name: contrato; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.contrato (
    cont_id integer NOT NULL,
    cont_descricao character varying(100) NOT NULL,
    cont_valor numeric(15,4) NOT NULL,
    cont_dia_lancamento integer,
    cont_dia_vencimento integer
);


ALTER TABLE public.contrato OWNER TO postgres;

--
-- TOC entry 313 (class 1259 OID 25313)
-- Name: contrato_cliente; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.contrato_cliente (
    contc_id integer NOT NULL,
    contc_id_contrato integer NOT NULL,
    contc_id_cliente integer NOT NULL,
    contc_data_contratacao timestamp with time zone NOT NULL,
    contc_data_encerramento timestamp with time zone,
    contc_status character varying(1) NOT NULL,
    contc_id_terminal integer NOT NULL
);


ALTER TABLE public.contrato_cliente OWNER TO postgres;

--
-- TOC entry 312 (class 1259 OID 25312)
-- Name: contrato_cliente_contc_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.contrato_cliente ALTER COLUMN contc_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.contrato_cliente_contc_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 310 (class 1259 OID 25303)
-- Name: contrato_cont_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.contrato ALTER COLUMN cont_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.contrato_cont_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 276 (class 1259 OID 24985)
-- Name: crediario_parcela; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.crediario_parcela (
    crep_id integer NOT NULL,
    crep_numero_crediario integer NOT NULL,
    crep_numero integer NOT NULL,
    crep_numero_parcelas integer NOT NULL,
    crep_valor numeric(15,2) NOT NULL,
    crep_valor_juros numeric(15,2) DEFAULT 0.00,
    crep_valor_desconto numeric(15,2) DEFAULT 0.00,
    crep_valor_total numeric(15,2) NOT NULL,
    crep_data_vencimento date NOT NULL,
    crep_data_pagamento timestamp with time zone,
    crep_status character varying(1) NOT NULL
);


ALTER TABLE public.crediario_parcela OWNER TO postgres;

--
-- TOC entry 275 (class 1259 OID 24984)
-- Name: crediario_parcela_crep_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.crediario_parcela ALTER COLUMN crep_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.crediario_parcela_crep_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 224 (class 1259 OID 24624)
-- Name: empresa; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.empresa (
    emp_id integer NOT NULL,
    emp_razao_social character varying(255) NOT NULL,
    emp_fantasia character varying(255),
    emp_cpf_cnpj character varying(14) NOT NULL,
    emp_inscricao_estadual character varying(18),
    emp_cep character varying(8),
    emp_logradouro character varying(100),
    emp_numero character varying(15),
    emp_complemento character varying(20),
    emp_bairro character varying(80),
    emp_cidade character varying(80),
    emp_codigo_municipio integer,
    emp_uf character varying(2),
    emp_email character varying(255),
    emp_telefone_principal character varying(15),
    emp_telefone_secundario character varying(15),
    emp_logotipo bytea,
    emp_regime integer,
    emp_id_csc character varying(10),
    emp_csc character varying(50),
    emp_contador_nome character varying(100),
    emp_contador_cpf_cnpj character varying(14),
    emp_contador_crc character varying(30),
    emp_contador_email character varying(255),
    emp_contador_cep character varying(8),
    emp_contador_logradouro character varying(100),
    emp_contador_numero character varying(15),
    emp_contador_complemento character varying(20),
    emp_contador_bairro character varying(80),
    emp_contador_cidade character varying(80),
    emp_contador_codigo_municipio integer,
    emp_contador_uf character varying(2)
);


ALTER TABLE public.empresa OWNER TO postgres;

--
-- TOC entry 223 (class 1259 OID 24623)
-- Name: empresa_emp_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.empresa ALTER COLUMN emp_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.empresa_emp_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 354 (class 1259 OID 59166)
-- Name: estoque_movimento; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.estoque_movimento (
    estm_id integer NOT NULL,
    estm_id_empresa integer NOT NULL,
    estm_numero integer,
    estm_data timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    estm_tipo character varying(2) NOT NULL,
    estm_status character varying(1) NOT NULL,
    estm_id_funcionario integer,
    estm_id_nota_fiscal integer,
    estm_numero_pedido integer,
    estm_observacao character varying(255),
    CONSTRAINT estm_numero_check CHECK ((((estm_tipo)::text = ANY ((ARRAY['NF'::character varying, 'PD'::character varying])::text[])) = (estm_numero IS NULL))),
    CONSTRAINT estm_status_check CHECK (((estm_status)::text = ANY ((ARRAY['P'::character varying, 'C'::character varying, 'X'::character varying])::text[]))),
    CONSTRAINT estm_tipo_check CHECK (((estm_tipo)::text = ANY ((ARRAY['NF'::character varying, 'PD'::character varying, 'PR'::character varying, 'EA'::character varying, 'DP'::character varying, 'SB'::character varying, 'SI'::character varying])::text[])))
);


ALTER TABLE public.estoque_movimento OWNER TO postgres;

--
-- TOC entry 353 (class 1259 OID 59165)
-- Name: estoque_movimento_estm_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.estoque_movimento ALTER COLUMN estm_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.estoque_movimento_estm_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 356 (class 1259 OID 59205)
-- Name: estoque_movimento_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.estoque_movimento_item (
    estmi_id integer NOT NULL,
    estmi_id_estoque_movimento integer NOT NULL,
    estmi_numero integer NOT NULL,
    estmi_id_produto_empresa_grade_item integer CONSTRAINT estoque_movimento_item_estmi_id_produto_empresa_grade__not_null NOT NULL,
    estmi_natureza character varying(1) NOT NULL,
    estmi_tipo character varying(1) DEFAULT 'N'::character varying NOT NULL,
    estmi_quantidade_estoque numeric(15,4) DEFAULT 0,
    estmi_quantidade_prateleira numeric(15,4) DEFAULT 0,
    estmi_id_nota_fiscal_item integer,
    estmi_id_pedido_item integer,
    CONSTRAINT estmi_natureza_check CHECK (((estmi_natureza)::text = ANY ((ARRAY['E'::character varying, 'S'::character varying])::text[]))),
    CONSTRAINT estmi_quantidade_check CHECK (((COALESCE(estmi_quantidade_estoque, (0)::numeric) > (0)::numeric) OR (COALESCE(estmi_quantidade_prateleira, (0)::numeric) > (0)::numeric))),
    CONSTRAINT estmi_tipo_check CHECK (((estmi_tipo)::text = ANY ((ARRAY['N'::character varying, 'P'::character varying])::text[]))),
    CONSTRAINT estoque_movimento_item_estmi_quantidade_estoque_check CHECK ((estmi_quantidade_estoque >= (0)::numeric)),
    CONSTRAINT estoque_movimento_item_estmi_quantidade_prateleira_check CHECK ((estmi_quantidade_prateleira >= (0)::numeric))
);


ALTER TABLE public.estoque_movimento_item OWNER TO postgres;

--
-- TOC entry 355 (class 1259 OID 59204)
-- Name: estoque_movimento_item_estmi_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.estoque_movimento_item ALTER COLUMN estmi_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.estoque_movimento_item_estmi_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 226 (class 1259 OID 24641)
-- Name: funcionario; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.funcionario (
    fun_id integer NOT NULL,
    fun_nome character varying(100) NOT NULL,
    fun_apelido character varying(40),
    fun_cpf character varying(11),
    fun_pis character varying(11),
    fun_telefone_principal character varying(15),
    fun_telefone_secundario character varying(15),
    fun_email character varying(255),
    fun_cep character varying(8),
    fun_logradouro character varying(100),
    fun_numero character varying(15),
    fun_complemento character varying(20),
    fun_bairro character varying(80),
    fun_cidade character varying(80),
    fun_codigo_municipio integer,
    fun_uf character varying(2),
    fun_usuario character varying(30),
    fun_senha character varying(60),
    fun_id_funcionario_cargo integer NOT NULL,
    fun_tipo character varying(1) NOT NULL
);


ALTER TABLE public.funcionario OWNER TO postgres;

--
-- TOC entry 235 (class 1259 OID 24696)
-- Name: funcionario_cargo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.funcionario_cargo (
    func_id integer NOT NULL,
    func_descricao character varying(60) NOT NULL
);


ALTER TABLE public.funcionario_cargo OWNER TO postgres;

--
-- TOC entry 234 (class 1259 OID 24695)
-- Name: funcionario_cargo_func_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.funcionario_cargo ALTER COLUMN func_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.funcionario_cargo_func_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 236 (class 1259 OID 24703)
-- Name: funcionario_cargo_permissao_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.funcionario_cargo_permissao_item (
    fcpi_id_funcionario_cargo integer CONSTRAINT funcionario_cargo_permissao__fcpi_id_funcionario_cargo_not_null NOT NULL,
    fcpi_chave_permissao_item character varying(60) CONSTRAINT funcionario_cargo_permissao__fcpi_chave_permissao_item_not_null NOT NULL
);


ALTER TABLE public.funcionario_cargo_permissao_item OWNER TO postgres;

--
-- TOC entry 225 (class 1259 OID 24640)
-- Name: funcionario_fun_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.funcionario ALTER COLUMN fun_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.funcionario_fun_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 266 (class 1259 OID 24906)
-- Name: grade; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.grade (
    gra_id integer NOT NULL,
    gra_descricao character varying(50) NOT NULL
);


ALTER TABLE public.grade OWNER TO postgres;

--
-- TOC entry 265 (class 1259 OID 24905)
-- Name: grade_gra_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.grade ALTER COLUMN gra_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.grade_gra_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 268 (class 1259 OID 24914)
-- Name: grade_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.grade_item (
    grai_id integer NOT NULL,
    grai_descricao character varying(30) NOT NULL,
    grai_id_grade integer NOT NULL
);


ALTER TABLE public.grade_item OWNER TO postgres;

--
-- TOC entry 267 (class 1259 OID 24913)
-- Name: grade_item_grai_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.grade_item ALTER COLUMN grai_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.grade_item_grai_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 279 (class 1259 OID 25016)
-- Name: ncm; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.ncm (
    ncm_codigo character varying(8) NOT NULL,
    ncm_descricao character varying(2048)
);


ALTER TABLE public.ncm OWNER TO postgres;

--
-- TOC entry 335 (class 1259 OID 25488)
-- Name: ncm_tributo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.ncm_tributo (
    ncmt_codigo_ncm character varying(8) NOT NULL,
    ncmt_uf character varying(2) NOT NULL,
    ncmt_percentual_aliquota_federal_nacional numeric(15,4) DEFAULT 0.00,
    ncmt_percentual_aliquota_federal_importado numeric(15,4) DEFAULT 0.00,
    ncmt_percentual_aliquota_estadual numeric(15,4) DEFAULT 0.00,
    ncmt_percentual_aliquota_municipal numeric(15,4) DEFAULT 0.00
);


ALTER TABLE public.ncm_tributo OWNER TO postgres;

--
-- TOC entry 285 (class 1259 OID 25054)
-- Name: nota_fiscal; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.nota_fiscal (
    not_id integer NOT NULL,
    not_numero integer NOT NULL,
    not_codigo character varying(8),
    not_serie integer NOT NULL,
    not_modelo character varying(3) NOT NULL,
    not_finalidade integer NOT NULL,
    not_chave character varying(44),
    not_data_lancamento timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    not_data_emissao timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    not_data_saida timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    not_valor_subtotal numeric(15,2) NOT NULL,
    not_valor_icms numeric(15,2) DEFAULT 0.00,
    not_valor_base_calculo_icms numeric(15,2) DEFAULT 0.00,
    not_valor_icms_desonerado numeric(15,2) DEFAULT 0.00,
    not_valor_st numeric(15,2) DEFAULT 0.00,
    not_valor_base_calculo_st numeric(15,2) DEFAULT 0.00,
    not_valor_frete numeric(15,2) DEFAULT 0.00,
    not_valor_seguro numeric(15,2) DEFAULT 0.00,
    not_valor_desconto numeric(15,2) DEFAULT 0.00,
    not_valor_acrescimo numeric(15,2) DEFAULT 0.00,
    not_valor_ipi numeric(15,2) DEFAULT 0.00,
    not_valor_pis numeric(15,2) DEFAULT 0.00,
    not_valor_cofins numeric(15,2) DEFAULT 0.00,
    not_valor_fcp numeric(15,2) DEFAULT 0.00,
    not_valor_total numeric(15,2) DEFAULT 0.00,
    not_observacao character varying(255),
    not_consumidor_final boolean NOT NULL,
    not_situacao character varying(2) NOT NULL,
    not_status character varying(1) NOT NULL,
    not_tipo character varying(1) NOT NULL,
    not_data_inclusao timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    not_justificativa character varying(255),
    not_transporte_modalidade_frete integer,
    not_transporte_transportador_razao_social character varying(60),
    not_transporte_transportador_inscricao_estadual character varying(14),
    not_transporte_transportador_cpf_cnpj character varying(14),
    not_transporte_transportador_bairro character varying(80),
    not_transporte_transportador_complemento character varying(20),
    not_transporte_transportador_uf character varying(2),
    not_transporte_transportador_numero character varying(15),
    not_transporte_transportador_logradouro character varying(100),
    not_transporte_transportador_cep character varying(8),
    not_transporte_transportador_codigo_municipio integer,
    not_transporte_transportador_cidade character varying(80),
    not_codigo_cfop character varying(4),
    not_id_empresa integer NOT NULL,
    not_id_funcionario integer NOT NULL,
    not_id_cliente integer,
    not_id_fornecedor integer,
    not_id_nota_fiscal_contingencia integer,
    not_arquivo xml,
    not_valor_base_calculo_imposto_seletivo numeric(15,2) DEFAULT 0.00,
    not_valor_imposto_seletivo numeric(15,2) DEFAULT 0.00,
    not_valor_base_calculo_ibs_cbs numeric(15,2) DEFAULT 0.00,
    not_valor_diferimento_uf numeric(15,2) DEFAULT 0.00,
    not_valor_devolucao_tributos_uf numeric(15,2) DEFAULT 0.00,
    not_valor_ibs_uf numeric(15,2) DEFAULT 0.00,
    not_valor_diferimento_municipio numeric(15,2) DEFAULT 0.00,
    not_valor_devolucao_tributos_municipio numeric(15,2) DEFAULT 0.00,
    not_valor_ibs_municipio numeric(15,2) DEFAULT 0.00,
    not_valor_ibs numeric(15,2) DEFAULT 0.00,
    not_valor_credito_presumido_municipio numeric(15,2) DEFAULT 0.00,
    not_valor_credito_presumido_condicao_suspensiva_municipio numeric(15,2) DEFAULT 0.00,
    not_valor_credito_presumido_cbs numeric(15,2) DEFAULT 0.00,
    not_valor_credito_presumido_condicao_suspensiva_cbs numeric(15,2) DEFAULT 0.00,
    not_valor_diferimento_cbs numeric(15,2) DEFAULT 0.00,
    not_valor_devolucao_tributos_cbs numeric(15,2) DEFAULT 0.00,
    not_valor_cbs numeric(15,2) DEFAULT 0.00,
    not_valor_base_calculo_valor_aproximado_tributo numeric(15,2) DEFAULT 0.00,
    not_valor_aproximado_tributo_municipal numeric(15,2) DEFAULT 0.00,
    not_valor_aproximado_tributo_estadual numeric(15,2) DEFAULT 0.00,
    not_valor_aproximado_tributo_federal numeric(15,2) DEFAULT 0.00,
    not_id_terminal integer
);


ALTER TABLE public.nota_fiscal OWNER TO postgres;

--
-- TOC entry 300 (class 1259 OID 25226)
-- Name: nota_fiscal_evento; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.nota_fiscal_evento (
    notfe_id integer NOT NULL,
    notfe_id_nota_fiscal integer,
    notfe_codigo character varying(2),
    notfe_arquivo xml,
    notfe_data timestamp with time zone DEFAULT CURRENT_TIMESTAMP
);


ALTER TABLE public.nota_fiscal_evento OWNER TO postgres;

--
-- TOC entry 299 (class 1259 OID 25225)
-- Name: nota_fiscal_evento_notfe_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.nota_fiscal_evento ALTER COLUMN notfe_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.nota_fiscal_evento_notfe_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 291 (class 1259 OID 25134)
-- Name: nota_fiscal_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.nota_fiscal_item (
    notfi_id integer NOT NULL,
    notfi_id_nota_fiscal integer NOT NULL,
    notfi_id_produto_empresa_grade_item integer NOT NULL,
    notfi_numero integer NOT NULL,
    notfi_codigo_cfop character varying(4) NOT NULL,
    notfi_ncm character varying(8) NOT NULL,
    notfi_cest character varying(7) NOT NULL,
    notfi_quantidade numeric(15,4) NOT NULL,
    notfi_valor numeric(15,4) NOT NULL,
    notfi_valor_acrescimo numeric(15,2) DEFAULT 0.00,
    notfi_valor_desconto numeric(15,2) DEFAULT 0.00,
    notfi_medida character varying(6) NOT NULL,
    notfi_origem character varying(1) NOT NULL,
    notfi_csosn character varying(3),
    notfi_cst character varying(2),
    notfi_valor_icms numeric(15,2) DEFAULT 0.00,
    notfi_valor_base_calculo_icms numeric(15,2) DEFAULT 0.00,
    notfi_percentual_icms numeric(7,4) DEFAULT 0.00,
    notfi_valor_icms_desonerado numeric(15,2) DEFAULT 0.00,
    notfi_motivo_icms_desonerado integer,
    notfi_deduz_icms_desonerado integer,
    notfi_valor_st numeric(15,2) DEFAULT 0.00,
    notfi_valor_base_calculo_st numeric(15,2) DEFAULT 0.00,
    notfi_percentual_reducao_base_calculo numeric(15,2) DEFAULT 0.00,
    notfi_percentual_st numeric(7,4) DEFAULT 0.00,
    notfi_valor_frete numeric(15,2) DEFAULT 0.00,
    notfi_valor_seguro numeric(15,2) DEFAULT 0.00,
    notfi_medida_tributavel character varying(6) NOT NULL,
    notfi_valor_total numeric(15,2) NOT NULL,
    notfi_numero_pedido character varying(15),
    notfi_numero_pedido_item character varying(10),
    notfi_valor_base_calculo_fcp numeric(15,2),
    notfi_percentual_fcp numeric(15,4),
    notfi_valor_fcp numeric(15,2),
    notfi_percentual_mva numeric(15,4),
    notfi_situacao_tributaria_pis_cofins character varying(2),
    notfi_valor_base_calculo_pis numeric(15,2),
    notfi_percentual_pis numeric(15,4) DEFAULT 0.00,
    notfi_valor_pis numeric(15,2),
    notfi_valor_base_calculo_cofins numeric(15,2),
    notfi_percentual_cofins numeric(15,4) DEFAULT 0.00,
    notfi_valor_cofins numeric(15,2),
    notfi_valor_base_calculo_ipi numeric(15,2),
    notfi_percentual_ipi numeric(15,2) DEFAULT 0.00,
    notfi_valor_ipi numeric(15,2),
    notfi_situacao_tributaria_ipi character varying(2),
    notfi_cst_ibs_cbs character varying(3),
    notfi_classificacao_tributaria_ibs_cbs character varying(6),
    notfi_valor_base_calculo_ibs_cbs numeric(15,2) DEFAULT 0.00,
    notfi_percentual_ibs_uf numeric(15,4) DEFAULT 0.00,
    notfi_valor_ibs_uf numeric(15,2) DEFAULT 0.00,
    notfi_percentual_ibs_municipio numeric(15,4) DEFAULT 0.00,
    notfi_valor_ibs_municipio numeric(15,2) DEFAULT 0.00,
    notfi_percentual_cbs numeric(15,4) DEFAULT 0.00,
    notfi_valor_cbs numeric(15,2) DEFAULT 0.00,
    notfi_valor_base_calculo_valor_aproximado_tributo numeric(15,2) DEFAULT 0.00,
    notfi_percentual_valor_aproximado_tributo_municipal numeric(15,4) DEFAULT 0.00,
    notfi_valor_aproximado_tributo_municipal numeric(15,2) DEFAULT 0.00,
    notfi_percentual_valor_aproximado_tributo_estadual numeric(15,4) DEFAULT 0.00,
    notfi_valor_aproximado_tributo_estadual numeric(15,2) DEFAULT 0.00,
    notfi_percentual_valor_aproximado_tributo_federal numeric(15,4) DEFAULT 0.00,
    notfi_valor_aproximado_tributo_federal numeric(15,2) DEFAULT 0.00,
    notfi_quantidade_estoque numeric(15,4) DEFAULT 0,
    notfi_quantidade_prateleira numeric(15,4) DEFAULT 0,
    notfi_valor_produto numeric(15,2) DEFAULT 0.00
);


ALTER TABLE public.nota_fiscal_item OWNER TO postgres;

--
-- TOC entry 290 (class 1259 OID 25133)
-- Name: nota_fiscal_item_notfi_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.nota_fiscal_item ALTER COLUMN notfi_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.nota_fiscal_item_notfi_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 284 (class 1259 OID 25053)
-- Name: nota_fiscal_not_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.nota_fiscal ALTER COLUMN not_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.nota_fiscal_not_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 289 (class 1259 OID 25123)
-- Name: nota_fiscal_pagamento; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.nota_fiscal_pagamento (
    notfpa_id integer NOT NULL,
    notfpa_meio_pagamento character varying(2) NOT NULL,
    notfpa_id_nota_fiscal integer NOT NULL,
    notfpa_tipo_pagamento integer NOT NULL,
    notfpa_valor numeric(15,2) NOT NULL,
    notfpa_tipo_integracao integer,
    notfpa_bandeira character varying(2),
    notfpa_codigo_autorizacao character varying(20),
    notfpa_valor_troco numeric(15,2),
    notfpa_cnpj_instituicao_pagamento character varying(14)
);


ALTER TABLE public.nota_fiscal_pagamento OWNER TO postgres;

--
-- TOC entry 288 (class 1259 OID 25122)
-- Name: nota_fiscal_pagamento_notfpa_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.nota_fiscal_pagamento ALTER COLUMN notfpa_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.nota_fiscal_pagamento_notfpa_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 287 (class 1259 OID 25114)
-- Name: nota_fiscal_parcela; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.nota_fiscal_parcela (
    notfp_id integer NOT NULL,
    notfp_id_nota_fiscal integer,
    notfp_numero character varying(3),
    notfp_data_vencimento timestamp with time zone NOT NULL,
    notfp_valor numeric(15,2) NOT NULL
);


ALTER TABLE public.nota_fiscal_parcela OWNER TO postgres;

--
-- TOC entry 286 (class 1259 OID 25113)
-- Name: nota_fiscal_parcela_notfp_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.nota_fiscal_parcela ALTER COLUMN notfp_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.nota_fiscal_parcela_notfp_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 298 (class 1259 OID 25218)
-- Name: nota_fiscal_referencia; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.nota_fiscal_referencia (
    notfr_id_nota_fiscal integer NOT NULL,
    notfr_id_nota_fiscal_referencia integer NOT NULL
);


ALTER TABLE public.nota_fiscal_referencia OWNER TO postgres;

--
-- TOC entry 293 (class 1259 OID 25182)
-- Name: nota_fiscal_transporte_volume; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.nota_fiscal_transporte_volume (
    notft_id integer NOT NULL,
    notft_id_nota_fiscal integer,
    notft_peso_liquido numeric(15,3),
    notft_marca character varying(60),
    notft_quantidade integer,
    notft_peso_bruto numeric(15,3),
    notft_especie character varying(60)
);


ALTER TABLE public.nota_fiscal_transporte_volume OWNER TO postgres;

--
-- TOC entry 292 (class 1259 OID 25181)
-- Name: nota_fiscal_transporte_volume_notft_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.nota_fiscal_transporte_volume ALTER COLUMN notft_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.nota_fiscal_transporte_volume_notft_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 295 (class 1259 OID 25189)
-- Name: pagamento; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.pagamento (
    pag_id integer NOT NULL,
    pag_id_terminal integer NOT NULL,
    pag_id_funcionario integer NOT NULL,
    pag_valor numeric(15,2) NOT NULL,
    pag_valor_troco numeric(15,2) DEFAULT 0,
    pag_valor_restante numeric(15,2) DEFAULT 0,
    pag_valor_pendente numeric(15,2) DEFAULT 0,
    pag_valor_confirmado numeric(15,2) DEFAULT 0,
    pag_valor_cancelado numeric(15,2) DEFAULT 0,
    pag_valor_credito numeric(15,2) DEFAULT 0.00,
    pag_data timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    pag_id_cliente integer,
    pag_natureza character varying(1) NOT NULL,
    pag_tipo character varying(1) DEFAULT 'V'::character varying NOT NULL,
    pag_id_terminal_historico integer,
    CONSTRAINT pag_tipo_check CHECK (((pag_tipo)::text = ANY ((ARRAY['V'::character varying, 'R'::character varying, 'A'::character varying, 'S'::character varying, 'U'::character varying, 'O'::character varying])::text[])))
);


ALTER TABLE public.pagamento OWNER TO postgres;

--
-- TOC entry 297 (class 1259 OID 25207)
-- Name: pagamento_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.pagamento_item (
    pagi_id integer NOT NULL,
    pagi_id_pagamento integer NOT NULL,
    pagi_forma_pagamento character varying(2) NOT NULL,
    pagi_valor numeric(15,2),
    pagi_status character varying(1) NOT NULL,
    pagi_id_conta_bancaria integer,
    pagi_transacao_identificador character varying(255),
    pagi_transacao_referencia character varying(255),
    pagi_transacao_adquirente character varying(50)
);


ALTER TABLE public.pagamento_item OWNER TO postgres;

--
-- TOC entry 296 (class 1259 OID 25206)
-- Name: pagamento_item_pagi_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.pagamento_item ALTER COLUMN pagi_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.pagamento_item_pagi_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 294 (class 1259 OID 25188)
-- Name: pagamento_pag_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.pagamento ALTER COLUMN pag_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.pagamento_pag_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 240 (class 1259 OID 24738)
-- Name: pais; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.pais (
    pai_codigo character varying(4) NOT NULL,
    pai_descricao character varying(60) NOT NULL
);


ALTER TABLE public.pais OWNER TO postgres;

--
-- TOC entry 228 (class 1259 OID 24663)
-- Name: parametro; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.parametro (
    par_chave character varying(60) NOT NULL,
    par_descricao character varying(255) NOT NULL
);


ALTER TABLE public.parametro OWNER TO postgres;

--
-- TOC entry 237 (class 1259 OID 24710)
-- Name: parametro_empresa; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.parametro_empresa (
    pare_id_empresa integer NOT NULL,
    pare_chave_parametro character varying(60) NOT NULL,
    pare_valor character varying(255)
);


ALTER TABLE public.parametro_empresa OWNER TO postgres;

--
-- TOC entry 272 (class 1259 OID 24944)
-- Name: pedido; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.pedido (
    ped_numero integer NOT NULL,
    ped_id_terminal integer NOT NULL,
    ped_id_cliente integer,
    ped_id_funcionario integer,
    ped_id_funcionario_cancelamento integer,
    ped_data timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    ped_data_fechamento timestamp with time zone,
    ped_valor_acrescimo numeric(15,2) DEFAULT 0.00,
    ped_valor_desconto numeric(15,2) DEFAULT 0.00,
    ped_valor_frete numeric(15,2) DEFAULT 0.00,
    ped_valor_subtotal numeric(15,2) NOT NULL,
    ped_valor_total numeric(15,2) NOT NULL,
    ped_quantidade_itens integer DEFAULT 0,
    ped_tipo character varying(1) NOT NULL,
    ped_status character varying(1) NOT NULL,
    ped_observacao character varying(255),
    cred_numero_crediario_refinanciado integer,
    cred_id_contrato_cliente integer,
    cred_valor_juros numeric(15,2) DEFAULT 0.00,
    cred_valor_pagamento_previo numeric(15,2),
    ped_id_nota_fiscal integer,
    ped_endereco_entrega_cep character varying(8),
    ped_endereco_entrega_logradouro character varying(100),
    ped_endereco_entrega_numero character varying(15),
    ped_endereco_entrega_uf character varying(2),
    ped_endereco_entrega_cidade character varying(80),
    ped_endereco_entrega_complemento character varying(20),
    ped_endereco_entrega_bairro character varying(80),
    ped_endereco_entrega_referencia character varying(255),
    ven_id_comanda integer,
    ped_justificativa_cancelamento character varying(255),
    ped_data_cancelamento timestamp with time zone,
    ped_id_terminal_historico integer
);


ALTER TABLE public.pedido OWNER TO postgres;

--
-- TOC entry 274 (class 1259 OID 24964)
-- Name: pedido_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.pedido_item (
    pedi_id integer NOT NULL,
    pedi_numero_pedido integer NOT NULL,
    pedi_id_produto_empresa_grade_item integer NOT NULL,
    pedi_numero integer NOT NULL,
    pedi_id_funcionario integer NOT NULL,
    pedi_id_funcionario_cancelamento integer,
    pedi_valor numeric(15,4) NOT NULL,
    pedi_quantidade numeric(15,4) NOT NULL,
    pedi_valor_desconto numeric(15,2) NOT NULL,
    pedi_valor_acrescimo numeric(15,2) NOT NULL,
    pedi_valor_subtotal numeric(15,2) NOT NULL,
    pedi_valor_total numeric(15,2) NOT NULL,
    pedi_data timestamp with time zone DEFAULT CURRENT_TIMESTAMP,
    pedi_status character varying(1) NOT NULL,
    pedi_observacao character varying(255),
    pedi_justificativa_cancelamento character varying(255),
    pedi_data_cancelamento timestamp with time zone,
    pedi_valor_adicional numeric(15,4) DEFAULT 0.00 NOT NULL
);


ALTER TABLE public.pedido_item OWNER TO postgres;

--
-- TOC entry 352 (class 1259 OID 59096)
-- Name: pedido_item_opcao; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.pedido_item_opcao (
    pedio_id integer NOT NULL,
    pedio_id_pedido_item integer NOT NULL,
    pedio_id_produto_empresa_grade_item integer,
    pedio_descricao_opcao character varying(50) NOT NULL,
    pedio_descricao_opcao_item character varying(50) NOT NULL,
    pedio_tipo character varying(1) NOT NULL,
    pedio_valor numeric(15,4) DEFAULT 0.00,
    pedio_quantidade numeric(15,4) DEFAULT 1.0000,
    pedio_ordem integer DEFAULT 0,
    CONSTRAINT pedio_tipo_check CHECK (((pedio_tipo)::text = ANY ((ARRAY['A'::character varying, 'O'::character varying])::text[])))
);


ALTER TABLE public.pedido_item_opcao OWNER TO postgres;

--
-- TOC entry 351 (class 1259 OID 59095)
-- Name: pedido_item_opcao_pedio_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.pedido_item_opcao ALTER COLUMN pedio_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.pedido_item_opcao_pedio_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 273 (class 1259 OID 24963)
-- Name: pedido_item_pedi_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.pedido_item ALTER COLUMN pedi_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.pedido_item_pedi_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 278 (class 1259 OID 25003)
-- Name: pedido_pagamento; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.pedido_pagamento (
    pedpg_id integer NOT NULL,
    pedpg_numero_pedido integer NOT NULL,
    pedpg_id_pagamento integer NOT NULL,
    pedpg_valor numeric(15,2) NOT NULL,
    pedpg_status character varying(1) NOT NULL,
    pedpg_data timestamp with time zone DEFAULT CURRENT_TIMESTAMP
);


ALTER TABLE public.pedido_pagamento OWNER TO postgres;

--
-- TOC entry 277 (class 1259 OID 25002)
-- Name: pedido_pagamento_pedpg_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.pedido_pagamento ALTER COLUMN pedpg_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.pedido_pagamento_pedpg_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 271 (class 1259 OID 24943)
-- Name: pedido_ped_numero_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.pedido ALTER COLUMN ped_numero ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.pedido_ped_numero_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 233 (class 1259 OID 24687)
-- Name: permissao_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.permissao_item (
    peri_chave character varying(60) NOT NULL,
    peri_descricao character varying(255) NOT NULL,
    peri_id_permissao_submodulo integer NOT NULL
);


ALTER TABLE public.permissao_item OWNER TO postgres;

--
-- TOC entry 230 (class 1259 OID 24671)
-- Name: permissao_modulo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.permissao_modulo (
    perm_id integer NOT NULL,
    perm_descricao character varying(255) NOT NULL
);


ALTER TABLE public.permissao_modulo OWNER TO postgres;

--
-- TOC entry 229 (class 1259 OID 24670)
-- Name: permissao_modulo_perm_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.permissao_modulo ALTER COLUMN perm_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.permissao_modulo_perm_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 232 (class 1259 OID 24679)
-- Name: permissao_submodulo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.permissao_submodulo (
    pers_id integer NOT NULL,
    pers_descricao character varying(255) NOT NULL,
    pers_id_permissao_modulo integer NOT NULL
);


ALTER TABLE public.permissao_submodulo OWNER TO postgres;

--
-- TOC entry 231 (class 1259 OID 24678)
-- Name: permissao_submodulo_pers_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.permissao_submodulo ALTER COLUMN pers_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.permissao_submodulo_pers_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 250 (class 1259 OID 24784)
-- Name: produto; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto (
    pro_id integer NOT NULL,
    pro_codigo character varying(20) NOT NULL,
    pro_codigo_barras character varying(20),
    pro_codigo_fornecedor character varying(30),
    pro_referencia character varying(32),
    pro_descricao character varying(100) NOT NULL,
    pro_descricao_resumida character varying(30),
    pro_fator numeric(15,4) DEFAULT 1,
    pro_condicao_pagamento integer DEFAULT 1,
    pro_peso numeric(10,2) DEFAULT 0,
    pro_balanca boolean DEFAULT false,
    pro_validade integer,
    pro_observacao character varying(255),
    pro_destacar boolean DEFAULT false,
    pro_imagem bytea,
    pro_codigo_produto_medida character varying(6) NOT NULL,
    pro_id_produto_marca integer,
    pro_codigo_produto_tipo character varying(2) NOT NULL,
    pro_id_produto_localizacao integer,
    pro_id_produto_grupo integer,
    pro_id_produto_subgrupo integer,
    pro_id_produto_linha integer,
    pro_id_grade integer NOT NULL,
    pro_origem character varying(1),
    pro_cst character varying(2),
    pro_csosn character varying(3),
    pro_codigo_ncm character varying(8),
    pro_cest character varying(7),
    pro_id_classificacao_tributaria_item integer
);


ALTER TABLE public.produto OWNER TO postgres;

--
-- TOC entry 252 (class 1259 OID 24805)
-- Name: produto_empresa; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_empresa (
    proe_id integer NOT NULL,
    proe_id_produto integer NOT NULL,
    proe_id_empresa integer NOT NULL,
    proe_estoque_minimo numeric(15,4),
    proe_percentual_reducao_base_calculo numeric(15,4) DEFAULT 0.00,
    proe_situacao_tributaria_pis_cofins_entrada character varying(2),
    proe_situacao_tributaria_pis_cofins_saida character varying(2),
    proe_situacao_tributaria_ipi_entrada character varying(2),
    proe_situacao_tributaria_ipi_saida character varying(2),
    proe_ativo boolean DEFAULT true,
    proe_percentual_pis numeric(15,4) DEFAULT 0.00,
    proe_percentual_cofins numeric(15,4) DEFAULT 0.00,
    proe_percentual_mva_original numeric(15,4) DEFAULT 0.00,
    proe_percentual_ipi numeric(15,4),
    proe_id_uf_produto integer,
    proe_codigo_produto_empresa_departamento_fiscal character varying(2),
    proe_valor_custo numeric(15,4) DEFAULT 0,
    proe_custo_percentual_ipi numeric(15,4) DEFAULT 0.00,
    proe_custo_aliquota_interna numeric(15,2) DEFAULT 0.00,
    proe_custo_aliquota_interestadual numeric(15,2) DEFAULT 0.00,
    proe_custo_percentual_fcp numeric(15,4) DEFAULT 0.00,
    proe_custo_percentual_despesa_fixa numeric(15,2) DEFAULT 0.00,
    proe_custo_percentual_despesa_variavel numeric(15,2) DEFAULT 0.00,
    proe_valor_custo_bruto numeric(15,4) DEFAULT 0,
    proe_custo_percentual_entrega numeric(15,2) DEFAULT 0.00,
    proe_custo_percentual_perda numeric(15,2) DEFAULT 0.00,
    proe_custo_percentual_gratificacao numeric(15,2) DEFAULT 0.00,
    proe_valor_compra_final numeric(15,4) DEFAULT 0,
    CONSTRAINT produto_empresa_proe_estoque_minimo_check CHECK ((proe_estoque_minimo >= (0)::numeric))
);


ALTER TABLE public.produto_empresa OWNER TO postgres;

--
-- TOC entry 317 (class 1259 OID 25344)
-- Name: produto_empresa_departamento_fiscal; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_empresa_departamento_fiscal (
    proedf_codigo character varying(2) NOT NULL,
    proedf_descricao character varying(60),
    proedf_valor numeric(15,4)
);


ALTER TABLE public.produto_empresa_departamento_fiscal OWNER TO postgres;

--
-- TOC entry 270 (class 1259 OID 24923)
-- Name: produto_empresa_grade_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_empresa_grade_item (
    proegi_id integer NOT NULL,
    proegi_quantidade_prateleira numeric(15,4) DEFAULT 0,
    proegi_quantidade_estoque numeric(15,4) DEFAULT 0,
    proegi_id_produto_empresa integer NOT NULL,
    proegi_id_grade_item integer NOT NULL,
    proegi_valor_custo_base numeric(15,4) DEFAULT 0,
    proegi_valor_custo_desconto numeric(15,4) DEFAULT 0,
    proegi_valor_custo_acrescimo numeric(15,4) DEFAULT 0,
    proegi_valor_custo numeric(15,4) DEFAULT 0,
    proegi_valor_venda numeric(15,4) DEFAULT 0,
    proegi_valor_venda_minimo numeric(15,4) DEFAULT 0,
    proegi_percentual_margem numeric(15,4) DEFAULT 0,
    proegi_valor_margem numeric(15,4) DEFAULT 0
);


ALTER TABLE public.produto_empresa_grade_item OWNER TO postgres;

--
-- TOC entry 302 (class 1259 OID 25236)
-- Name: produto_empresa_grade_item_fornecedor; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_empresa_grade_item_fornecedor (
    proegif_id integer NOT NULL,
    proegif_id_produto_empresa_grade_item integer CONSTRAINT produto_empresa_grade_item__proegif_id_produto_empresa_not_null NOT NULL,
    proegif_id_fornecedor integer CONSTRAINT produto_empresa_grade_item_forne_proegif_id_fornecedor_not_null NOT NULL,
    proegif_codigo character varying(60),
    proegif_codigo_barras character varying(20),
    proegif_fator numeric(15,4) DEFAULT 1
);


ALTER TABLE public.produto_empresa_grade_item_fornecedor OWNER TO postgres;

--
-- TOC entry 301 (class 1259 OID 25235)
-- Name: produto_empresa_grade_item_fornecedor_proegif_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_empresa_grade_item_fornecedor ALTER COLUMN proegif_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_empresa_grade_item_fornecedor_proegif_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 269 (class 1259 OID 24922)
-- Name: produto_empresa_grade_item_proegi_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_empresa_grade_item ALTER COLUMN proegi_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_empresa_grade_item_proegi_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 251 (class 1259 OID 24804)
-- Name: produto_empresa_proe_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_empresa ALTER COLUMN proe_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_empresa_proe_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 260 (class 1259 OID 24872)
-- Name: produto_grupo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_grupo (
    prog_id integer NOT NULL,
    prog_descricao character varying(50) NOT NULL,
    prog_imagem bytea
);


ALTER TABLE public.produto_grupo OWNER TO postgres;

--
-- TOC entry 259 (class 1259 OID 24871)
-- Name: produto_grupo_prog_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_grupo ALTER COLUMN prog_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_grupo_prog_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 264 (class 1259 OID 24895)
-- Name: produto_linha; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_linha (
    proli_id integer NOT NULL,
    proli_descricao character varying(50) NOT NULL,
    proli_id_subgrupo integer NOT NULL
);


ALTER TABLE public.produto_linha OWNER TO postgres;

--
-- TOC entry 263 (class 1259 OID 24894)
-- Name: produto_linha_proli_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_linha ALTER COLUMN proli_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_linha_proli_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 258 (class 1259 OID 24862)
-- Name: produto_localizacao; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_localizacao (
    prol_id integer NOT NULL,
    prol_descricao character varying(20) NOT NULL
);


ALTER TABLE public.produto_localizacao OWNER TO postgres;

--
-- TOC entry 257 (class 1259 OID 24861)
-- Name: produto_localizacao_prol_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_localizacao ALTER COLUMN prol_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_localizacao_prol_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 255 (class 1259 OID 24843)
-- Name: produto_marca; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_marca (
    prom_id integer NOT NULL,
    prom_descricao character varying(30) NOT NULL
);


ALTER TABLE public.produto_marca OWNER TO postgres;

--
-- TOC entry 254 (class 1259 OID 24842)
-- Name: produto_marca_prom_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_marca ALTER COLUMN prom_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_marca_prom_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 256 (class 1259 OID 24852)
-- Name: produto_medida; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_medida (
    prome_codigo character varying(6) NOT NULL,
    prome_descricao character varying(10) NOT NULL
);


ALTER TABLE public.produto_medida OWNER TO postgres;

--
-- TOC entry 346 (class 1259 OID 59064)
-- Name: produto_opcao; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_opcao (
    proo_id integer NOT NULL,
    proo_descricao character varying(50) NOT NULL,
    proo_tipo character varying(1) NOT NULL,
    proo_quantidade_minima integer DEFAULT 0 NOT NULL,
    proo_quantidade_maxima integer DEFAULT 1 NOT NULL,
    proo_ordem integer DEFAULT 0,
    proo_ativo boolean DEFAULT true,
    CONSTRAINT proo_quantidade_check CHECK (((proo_quantidade_maxima >= 1) AND (proo_quantidade_maxima >= proo_quantidade_minima))),
    CONSTRAINT proo_tipo_check CHECK (((proo_tipo)::text = ANY ((ARRAY['A'::character varying, 'O'::character varying])::text[])))
);


ALTER TABLE public.produto_opcao OWNER TO postgres;

--
-- TOC entry 348 (class 1259 OID 59077)
-- Name: produto_opcao_item; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_opcao_item (
    prooi_id integer NOT NULL,
    prooi_id_produto_opcao integer NOT NULL,
    prooi_descricao character varying(50) NOT NULL,
    prooi_valor numeric(15,4) DEFAULT 0.00,
    prooi_id_produto integer,
    prooi_id_grade_item integer,
    prooi_padrao boolean DEFAULT false,
    prooi_ordem integer DEFAULT 0,
    prooi_ativo boolean DEFAULT true
);


ALTER TABLE public.produto_opcao_item OWNER TO postgres;

--
-- TOC entry 347 (class 1259 OID 59076)
-- Name: produto_opcao_item_prooi_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_opcao_item ALTER COLUMN prooi_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_opcao_item_prooi_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 350 (class 1259 OID 59088)
-- Name: produto_opcao_produto; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_opcao_produto (
    proop_id integer NOT NULL,
    proop_id_produto integer NOT NULL,
    proop_id_produto_opcao integer NOT NULL,
    proop_quantidade_minima integer,
    proop_quantidade_maxima integer,
    proop_ordem integer DEFAULT 0,
    CONSTRAINT proop_quantidade_check CHECK (((proop_quantidade_maxima IS NULL) OR (proop_quantidade_maxima >= GREATEST(COALESCE(proop_quantidade_minima, 0), 1))))
);


ALTER TABLE public.produto_opcao_produto OWNER TO postgres;

--
-- TOC entry 349 (class 1259 OID 59087)
-- Name: produto_opcao_produto_proop_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_opcao_produto ALTER COLUMN proop_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_opcao_produto_proop_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 345 (class 1259 OID 59063)
-- Name: produto_opcao_proo_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_opcao ALTER COLUMN proo_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_opcao_proo_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 336 (class 1259 OID 25979)
-- Name: produto_pro_codigo_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

CREATE SEQUENCE public.produto_pro_codigo_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;


ALTER SEQUENCE public.produto_pro_codigo_seq OWNER TO postgres;

--
-- TOC entry 249 (class 1259 OID 24783)
-- Name: produto_pro_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto ALTER COLUMN pro_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_pro_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 262 (class 1259 OID 24884)
-- Name: produto_subgrupo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_subgrupo (
    pros_id integer NOT NULL,
    pros_descricao character varying(50) NOT NULL,
    pros_id_grupo integer NOT NULL
);


ALTER TABLE public.produto_subgrupo OWNER TO postgres;

--
-- TOC entry 261 (class 1259 OID 24883)
-- Name: produto_subgrupo_pros_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.produto_subgrupo ALTER COLUMN pros_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.produto_subgrupo_pros_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 253 (class 1259 OID 24833)
-- Name: produto_tipo; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.produto_tipo (
    prot_codigo character varying(2) NOT NULL,
    prot_descricao character varying(100) NOT NULL
);


ALTER TABLE public.produto_tipo OWNER TO postgres;

--
-- TOC entry 222 (class 1259 OID 24577)
-- Name: sistema; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.sistema (
    sis_id smallint DEFAULT 1 NOT NULL,
    sis_chave character varying(255),
    sis_versao character varying(10),
    sis_data_backup date,
    CONSTRAINT sistema_sis_id_check CHECK ((sis_id = 1))
);


ALTER TABLE public.sistema OWNER TO postgres;

--
-- TOC entry 283 (class 1259 OID 25041)
-- Name: situacao_tributaria; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.situacao_tributaria (
    sitt_id integer NOT NULL,
    sitt_codigo_cfop_origem character varying(4) NOT NULL,
    sitt_codigo_cfop_destino character varying(4) NOT NULL,
    sitt_codigo_produto_tipo character varying(2) NOT NULL,
    sitt_cst character varying(2) NOT NULL
);


ALTER TABLE public.situacao_tributaria OWNER TO postgres;

--
-- TOC entry 282 (class 1259 OID 25040)
-- Name: situacao_tributaria_sitt_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.situacao_tributaria ALTER COLUMN sitt_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.situacao_tributaria_sitt_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 307 (class 1259 OID 25265)
-- Name: terminal; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.terminal (
    ter_id integer NOT NULL,
    ter_descricao character varying(60) NOT NULL,
    ter_id_empresa integer NOT NULL,
    ter_numero_venda_pendente integer
);


ALTER TABLE public.terminal OWNER TO postgres;

--
-- TOC entry 344 (class 1259 OID 50644)
-- Name: terminal_historico_conferencia; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.terminal_historico_conferencia (
    terhc_id integer NOT NULL,
    terhc_id_terminal_historico integer CONSTRAINT terminal_historico_conferen_terhc_id_terminal_historic_not_null NOT NULL,
    terhc_forma_pagamento character varying(2) NOT NULL,
    terhc_valor_apurado numeric(15,2) DEFAULT 0.00,
    terhc_valor_conferido numeric(15,2) DEFAULT 0.00,
    terhc_valor_diferenca numeric(15,2) DEFAULT 0.00,
    terhc_justificativa character varying(255),
    terhc_id_funcionario integer,
    terhc_data_conferencia timestamp with time zone,
    CONSTRAINT terhc_forma_pagamento_check CHECK (((terhc_forma_pagamento)::text = ANY ((ARRAY['DN'::character varying, 'CH'::character varying, 'CC'::character varying, 'CD'::character varying, 'CL'::character varying, 'VA'::character varying, 'VR'::character varying, 'VP'::character varying, 'VC'::character varying, 'BL'::character varying, 'DB'::character varying, 'PX'::character varying, 'TB'::character varying, 'CB'::character varying, 'SP'::character varying, 'OT'::character varying])::text[])))
);


ALTER TABLE public.terminal_historico_conferencia OWNER TO postgres;

--
-- TOC entry 343 (class 1259 OID 50643)
-- Name: terminal_historico_conferencia_terhc_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.terminal_historico_conferencia ALTER COLUMN terhc_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.terminal_historico_conferencia_terhc_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 308 (class 1259 OID 25273)
-- Name: terminal_historico_terh_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.terminal_historico ALTER COLUMN terh_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.terminal_historico_terh_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 306 (class 1259 OID 25264)
-- Name: terminal_ter_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.terminal ALTER COLUMN ter_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.terminal_ter_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 315 (class 1259 OID 25325)
-- Name: uf_produto; Type: TABLE; Schema: public; Owner: postgres
--

CREATE TABLE public.uf_produto (
    ufp_id integer NOT NULL,
    ufp_uf character varying(2) NOT NULL,
    ufp_codigo_item_ipm character varying(100) NOT NULL,
    ufp_descricao character varying(512)
);


ALTER TABLE public.uf_produto OWNER TO postgres;

--
-- TOC entry 314 (class 1259 OID 25324)
-- Name: uf_produto_ufp_id_seq; Type: SEQUENCE; Schema: public; Owner: postgres
--

ALTER TABLE public.uf_produto ALTER COLUMN ufp_id ADD GENERATED BY DEFAULT AS IDENTITY (
    SEQUENCE NAME public.uf_produto_ufp_id_seq
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1
);


--
-- TOC entry 5539 (class 2606 OID 24662)
-- Name: acesso ace_id_funcionario_ace_id_empresa_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.acesso
    ADD CONSTRAINT ace_id_funcionario_ace_id_empresa_pkey PRIMARY KEY (ace_id_funcionario, ace_id_empresa);


--
-- TOC entry 5725 (class 2606 OID 25479)
-- Name: ajuda ajuda_codigo_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.ajuda
    ADD CONSTRAINT ajuda_codigo_pkey PRIMARY KEY (ajuda_codigo);


--
-- TOC entry 5698 (class 2606 OID 25343)
-- Name: aliquota aliq_uf_origem_aliq_uf_destino_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.aliquota
    ADD CONSTRAINT aliq_uf_origem_aliq_uf_destino_pkey PRIMARY KEY (aliq_uf_origem, aliq_uf_destino);


--
-- TOC entry 5706 (class 2606 OID 25399)
-- Name: boleto bol_id_crediario_parcela_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.boleto
    ADD CONSTRAINT bol_id_crediario_parcela_key UNIQUE (bol_id_crediario_parcela);


--
-- TOC entry 5708 (class 2606 OID 25397)
-- Name: boleto bol_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.boleto
    ADD CONSTRAINT bol_id_pkey PRIMARY KEY (bol_id);


--
-- TOC entry 5710 (class 2606 OID 25408)
-- Name: boleto_instrucao boli_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.boleto_instrucao
    ADD CONSTRAINT boli_id_pkey PRIMARY KEY (boli_id);


--
-- TOC entry 5642 (class 2606 OID 25031)
-- Name: cest cest_codigo_cest_ncm_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cest
    ADD CONSTRAINT cest_codigo_cest_ncm_pkey PRIMARY KEY (cest_ncm, cest_codigo);


--
-- TOC entry 5644 (class 2606 OID 25039)
-- Name: cfop cfop_codigo_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cfop
    ADD CONSTRAINT cfop_codigo_pkey PRIMARY KEY (cfop_codigo);


--
-- TOC entry 5501 (class 2606 OID 50677)
-- Name: cfop cfop_tipo_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.cfop
    ADD CONSTRAINT cfop_tipo_check CHECK (((cfop_tipo)::text = ANY ((ARRAY['E'::character varying, 'S'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5574 (class 2606 OID 24782)
-- Name: cliente_endereco clie_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_endereco
    ADD CONSTRAINT clie_id_pkey PRIMARY KEY (clie_id);


--
-- TOC entry 5555 (class 2606 OID 26069)
-- Name: cliente_fornecedor clif_codigo_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_fornecedor
    ADD CONSTRAINT clif_codigo_key UNIQUE (clif_codigo);


--
-- TOC entry 5557 (class 2606 OID 24737)
-- Name: cliente_fornecedor clif_cpf_cnpj_clif_inscricao_estadual_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_fornecedor
    ADD CONSTRAINT clif_cpf_cnpj_clif_inscricao_estadual_key UNIQUE (clif_cpf_cnpj, clif_inscricao_estadual);


--
-- TOC entry 5559 (class 2606 OID 24735)
-- Name: cliente_fornecedor clif_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_fornecedor
    ADD CONSTRAINT clif_id_pkey PRIMARY KEY (clif_id);


--
-- TOC entry 5495 (class 2606 OID 50676)
-- Name: cliente_fornecedor clif_tipo_cadastro_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.cliente_fornecedor
    ADD CONSTRAINT clif_tipo_cadastro_check CHECK (((clif_tipo_cadastro)::text = ANY ((ARRAY['C'::character varying, 'F'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5569 (class 2606 OID 24772)
-- Name: cliente_grupo clig_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_grupo
    ADD CONSTRAINT clig_descricao_key UNIQUE (clig_descricao);


--
-- TOC entry 5571 (class 2606 OID 24770)
-- Name: cliente_grupo clig_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_grupo
    ADD CONSTRAINT clig_id_pkey PRIMARY KEY (clig_id);


--
-- TOC entry 5563 (class 2606 OID 24754)
-- Name: cliente_rota clir_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_rota
    ADD CONSTRAINT clir_descricao_key UNIQUE (clir_descricao);


--
-- TOC entry 5565 (class 2606 OID 24752)
-- Name: cliente_rota clir_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_rota
    ADD CONSTRAINT clir_id_pkey PRIMARY KEY (clir_id);


--
-- TOC entry 5567 (class 2606 OID 24762)
-- Name: cliente_referencia_comercial clirc_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_referencia_comercial
    ADD CONSTRAINT clirc_id_pkey PRIMARY KEY (clirc_id);


--
-- TOC entry 5712 (class 2606 OID 25427)
-- Name: conta cnt_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta
    ADD CONSTRAINT cnt_id_pkey PRIMARY KEY (cnt_id);


--
-- TOC entry 5508 (class 2606 OID 50675)
-- Name: conta cnt_modalidade_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.conta
    ADD CONSTRAINT cnt_modalidade_check CHECK (((cnt_modalidade)::text = ANY ((ARRAY['P'::character varying, 'R'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5714 (class 2606 OID 25441)
-- Name: conta_parcela cntp_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_parcela
    ADD CONSTRAINT cntp_id_pkey PRIMARY KEY (cntp_id);


--
-- TOC entry 5716 (class 2606 OID 25455)
-- Name: conta_pagamento cntpg_id_conta_cntpg_id_pagamento_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_pagamento
    ADD CONSTRAINT cntpg_id_conta_cntpg_id_pagamento_key UNIQUE (cntpg_id_conta, cntpg_id_pagamento);


--
-- TOC entry 5719 (class 2606 OID 25453)
-- Name: conta_pagamento cntpg_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_pagamento
    ADD CONSTRAINT cntpg_id_pkey PRIMARY KEY (cntpg_id);


--
-- TOC entry 5723 (class 2606 OID 25470)
-- Name: conta_recorrencia cntr_id_conta_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_recorrencia
    ADD CONSTRAINT cntr_id_conta_pkey PRIMARY KEY (cntr_id_conta);


--
-- TOC entry 5721 (class 2606 OID 25463)
-- Name: conta_tipo cntt_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_tipo
    ADD CONSTRAINT cntt_id_pkey PRIMARY KEY (cntt_id);


--
-- TOC entry 5727 (class 2606 OID 34277)
-- Name: comanda com_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.comanda
    ADD CONSTRAINT com_id_pkey PRIMARY KEY (com_id);


--
-- TOC entry 5729 (class 2606 OID 34279)
-- Name: comanda com_numero_com_id_empresa_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.comanda
    ADD CONSTRAINT com_numero_com_id_empresa_key UNIQUE (com_numero, com_id_empresa);


--
-- TOC entry 5702 (class 2606 OID 25366)
-- Name: conta_bancaria conb_banco_conb_agencia_conb_numero_conb_digito_verificador_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_bancaria
    ADD CONSTRAINT conb_banco_conb_agencia_conb_numero_conb_digito_verificador_key UNIQUE (conb_agencia, conb_numero, conb_banco, conb_digito_verificador);


--
-- TOC entry 5704 (class 2606 OID 25364)
-- Name: conta_bancaria conb_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_bancaria
    ADD CONSTRAINT conb_id_pkey PRIMARY KEY (conb_id);


--
-- TOC entry 5733 (class 2606 OID 26043)
-- Name: conta_bancaria_retorno conbr_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_bancaria_retorno
    ADD CONSTRAINT conbr_id_pkey PRIMARY KEY (conbr_id);


--
-- TOC entry 5689 (class 2606 OID 25311)
-- Name: contrato cont_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.contrato
    ADD CONSTRAINT cont_id_pkey PRIMARY KEY (cont_id);


--
-- TOC entry 5692 (class 2606 OID 25323)
-- Name: contrato_cliente contc_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.contrato_cliente
    ADD CONSTRAINT contc_id_pkey PRIMARY KEY (contc_id);


--
-- TOC entry 5631 (class 2606 OID 24999)
-- Name: crediario_parcela crep_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.crediario_parcela
    ADD CONSTRAINT crep_id_pkey PRIMARY KEY (crep_id);


--
-- TOC entry 5633 (class 2606 OID 25001)
-- Name: crediario_parcela crep_numero_crediario_crep_numero_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.crediario_parcela
    ADD CONSTRAINT crep_numero_crediario_crep_numero_key UNIQUE (crep_numero_crediario, crep_numero);


--
-- TOC entry 5673 (class 2606 OID 25250)
-- Name: classificacao_tributaria ctrib_codigo_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.classificacao_tributaria
    ADD CONSTRAINT ctrib_codigo_pkey PRIMARY KEY (ctrib_codigo);


--
-- TOC entry 5675 (class 2606 OID 25263)
-- Name: classificacao_tributaria_item ctribi_codigo_ctribi_codigo_classificacao_tributaria_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.classificacao_tributaria_item
    ADD CONSTRAINT ctribi_codigo_ctribi_codigo_classificacao_tributaria_key UNIQUE (ctribi_codigo, ctribi_codigo_classificacao_tributaria);


--
-- TOC entry 5677 (class 2606 OID 25261)
-- Name: classificacao_tributaria_item ctribi_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.classificacao_tributaria_item
    ADD CONSTRAINT ctribi_id_pkey PRIMARY KEY (ctribi_id);


--
-- TOC entry 5525 (class 2606 OID 24635)
-- Name: empresa emp_cpf_cnpj_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.empresa
    ADD CONSTRAINT emp_cpf_cnpj_key UNIQUE (emp_cpf_cnpj);


--
-- TOC entry 5527 (class 2606 OID 24633)
-- Name: empresa emp_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.empresa
    ADD CONSTRAINT emp_id_pkey PRIMARY KEY (emp_id);


--
-- TOC entry 5529 (class 2606 OID 24637)
-- Name: empresa emp_inscricao_estadual_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.empresa
    ADD CONSTRAINT emp_inscricao_estadual_key UNIQUE (emp_inscricao_estadual);


--
-- TOC entry 5531 (class 2606 OID 24639)
-- Name: empresa emp_razao_social_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.empresa
    ADD CONSTRAINT emp_razao_social_key UNIQUE (emp_razao_social);


--
-- TOC entry 5757 (class 2606 OID 59177)
-- Name: estoque_movimento estm_id_empresa_estm_numero_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento
    ADD CONSTRAINT estm_id_empresa_estm_numero_key UNIQUE (estm_id_empresa, estm_numero);


--
-- TOC entry 5759 (class 2606 OID 59175)
-- Name: estoque_movimento estm_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento
    ADD CONSTRAINT estm_id_pkey PRIMARY KEY (estm_id);


--
-- TOC entry 5763 (class 2606 OID 59222)
-- Name: estoque_movimento_item estmi_id_estoque_movimento_estmi_numero_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento_item
    ADD CONSTRAINT estmi_id_estoque_movimento_estmi_numero_key UNIQUE (estmi_id_estoque_movimento, estmi_numero);


--
-- TOC entry 5765 (class 2606 OID 59220)
-- Name: estoque_movimento_item estmi_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento_item
    ADD CONSTRAINT estmi_id_pkey PRIMARY KEY (estmi_id);


--
-- TOC entry 5551 (class 2606 OID 24709)
-- Name: funcionario_cargo_permissao_item fcpi_id_funcionario_cargo_fcpi_chave_permissao_item_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.funcionario_cargo_permissao_item
    ADD CONSTRAINT fcpi_id_funcionario_cargo_fcpi_chave_permissao_item_pkey PRIMARY KEY (fcpi_id_funcionario_cargo, fcpi_chave_permissao_item);


--
-- TOC entry 5533 (class 2606 OID 24653)
-- Name: funcionario fun_cpf_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.funcionario
    ADD CONSTRAINT fun_cpf_key UNIQUE (fun_cpf);


--
-- TOC entry 5535 (class 2606 OID 24651)
-- Name: funcionario fun_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.funcionario
    ADD CONSTRAINT fun_id_pkey PRIMARY KEY (fun_id);


--
-- TOC entry 5494 (class 2606 OID 59051)
-- Name: funcionario fun_tipo_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.funcionario
    ADD CONSTRAINT fun_tipo_check CHECK (((fun_tipo)::text = ANY ((ARRAY['O'::character varying, 'A'::character varying, 'S'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5537 (class 2606 OID 24655)
-- Name: funcionario fun_usuario_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.funcionario
    ADD CONSTRAINT fun_usuario_key UNIQUE (fun_usuario);


--
-- TOC entry 5549 (class 2606 OID 24702)
-- Name: funcionario_cargo func_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.funcionario_cargo
    ADD CONSTRAINT func_id_pkey PRIMARY KEY (func_id);


--
-- TOC entry 5612 (class 2606 OID 24912)
-- Name: grade gra_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.grade
    ADD CONSTRAINT gra_id_pkey PRIMARY KEY (gra_id);


--
-- TOC entry 5614 (class 2606 OID 24921)
-- Name: grade_item grai_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.grade_item
    ADD CONSTRAINT grai_id_pkey PRIMARY KEY (grai_id);


--
-- TOC entry 5640 (class 2606 OID 25023)
-- Name: ncm ncm_codigo_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.ncm
    ADD CONSTRAINT ncm_codigo_pkey PRIMARY KEY (ncm_codigo);


--
-- TOC entry 5731 (class 2606 OID 25498)
-- Name: ncm_tributo ncmt_codigo_ncm_ncmt_uf_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.ncm_tributo
    ADD CONSTRAINT ncmt_codigo_ncm_ncmt_uf_pkey PRIMARY KEY (ncmt_codigo_ncm, ncmt_uf);


--
-- TOC entry 5650 (class 2606 OID 25112)
-- Name: nota_fiscal not_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal
    ADD CONSTRAINT not_id_pkey PRIMARY KEY (not_id);


--
-- TOC entry 5502 (class 2606 OID 50678)
-- Name: nota_fiscal not_tipo_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.nota_fiscal
    ADD CONSTRAINT not_tipo_check CHECK (((not_tipo)::text = ANY ((ARRAY['E'::character varying, 'S'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5667 (class 2606 OID 25224)
-- Name: nota_fiscal_referencia nota_fiscal_referencia_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_referencia
    ADD CONSTRAINT nota_fiscal_referencia_pkey PRIMARY KEY (notfr_id_nota_fiscal, notfr_id_nota_fiscal_referencia);


--
-- TOC entry 5669 (class 2606 OID 25234)
-- Name: nota_fiscal_evento notfe_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_evento
    ADD CONSTRAINT notfe_id_pkey PRIMARY KEY (notfe_id);


--
-- TOC entry 5656 (class 2606 OID 25180)
-- Name: nota_fiscal_item notfi_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_item
    ADD CONSTRAINT notfi_id_pkey PRIMARY KEY (notfi_id);


--
-- TOC entry 5652 (class 2606 OID 25121)
-- Name: nota_fiscal_parcela notfp_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_parcela
    ADD CONSTRAINT notfp_id_pkey PRIMARY KEY (notfp_id);


--
-- TOC entry 5654 (class 2606 OID 25132)
-- Name: nota_fiscal_pagamento notfpa_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_pagamento
    ADD CONSTRAINT notfpa_id_pkey PRIMARY KEY (notfpa_id);


--
-- TOC entry 5658 (class 2606 OID 25187)
-- Name: nota_fiscal_transporte_volume notft_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_transporte_volume
    ADD CONSTRAINT notft_id_pkey PRIMARY KEY (notft_id);


--
-- TOC entry 5661 (class 2606 OID 25205)
-- Name: pagamento pag_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pagamento
    ADD CONSTRAINT pag_id_pkey PRIMARY KEY (pag_id);


--
-- TOC entry 5503 (class 2606 OID 50668)
-- Name: pagamento pag_natureza_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.pagamento
    ADD CONSTRAINT pag_natureza_check CHECK (((pag_natureza)::text = ANY ((ARRAY['E'::character varying, 'S'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5505 (class 2606 OID 50670)
-- Name: pagamento_item pagi_forma_pagamento_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.pagamento_item
    ADD CONSTRAINT pagi_forma_pagamento_check CHECK (((pagi_forma_pagamento)::text = ANY ((ARRAY['DN'::character varying, 'CH'::character varying, 'CC'::character varying, 'CD'::character varying, 'CL'::character varying, 'VA'::character varying, 'VR'::character varying, 'VP'::character varying, 'VC'::character varying, 'BL'::character varying, 'DB'::character varying, 'PX'::character varying, 'TB'::character varying, 'CB'::character varying, 'SP'::character varying, 'OT'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5665 (class 2606 OID 25217)
-- Name: pagamento_item pagi_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pagamento_item
    ADD CONSTRAINT pagi_id_pkey PRIMARY KEY (pagi_id);


--
-- TOC entry 5506 (class 2606 OID 50669)
-- Name: pagamento_item pagi_status_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.pagamento_item
    ADD CONSTRAINT pagi_status_check CHECK (((pagi_status)::text = ANY ((ARRAY['P'::character varying, 'C'::character varying, 'X'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5561 (class 2606 OID 24744)
-- Name: pais pai_codigo_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pais
    ADD CONSTRAINT pai_codigo_pkey PRIMARY KEY (pai_codigo);


--
-- TOC entry 5541 (class 2606 OID 24669)
-- Name: parametro par_chave_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.parametro
    ADD CONSTRAINT par_chave_pkey PRIMARY KEY (par_chave);


--
-- TOC entry 5553 (class 2606 OID 24716)
-- Name: parametro_empresa pare_id_empresa_pare_chave_parametro_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.parametro_empresa
    ADD CONSTRAINT pare_id_empresa_pare_chave_parametro_pkey PRIMARY KEY (pare_id_empresa, pare_chave_parametro);


--
-- TOC entry 5623 (class 2606 OID 24962)
-- Name: pedido ped_numero_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT ped_numero_pkey PRIMARY KEY (ped_numero);


--
-- TOC entry 5497 (class 2606 OID 50672)
-- Name: pedido ped_status_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.pedido
    ADD CONSTRAINT ped_status_check CHECK (((ped_status)::text = ANY ((ARRAY['P'::character varying, 'S'::character varying, 'C'::character varying, 'X'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5498 (class 2606 OID 50671)
-- Name: pedido ped_tipo_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.pedido
    ADD CONSTRAINT ped_tipo_check CHECK (((ped_tipo)::text = ANY ((ARRAY['V'::character varying, 'C'::character varying, 'O'::character varying, 'D'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5626 (class 2606 OID 24981)
-- Name: pedido_item pedi_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item
    ADD CONSTRAINT pedi_id_pkey PRIMARY KEY (pedi_id);


--
-- TOC entry 5629 (class 2606 OID 24983)
-- Name: pedido_item pedi_numero_pedido_pedi_numero_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item
    ADD CONSTRAINT pedi_numero_pedido_pedi_numero_key UNIQUE (pedi_numero, pedi_numero_pedido);


--
-- TOC entry 5499 (class 2606 OID 50673)
-- Name: pedido_item pedi_status_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.pedido_item
    ADD CONSTRAINT pedi_status_check CHECK (((pedi_status)::text = ANY ((ARRAY['P'::character varying, 'C'::character varying, 'X'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5754 (class 2606 OID 59150)
-- Name: pedido_item_opcao pedio_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item_opcao
    ADD CONSTRAINT pedio_id_pkey PRIMARY KEY (pedio_id);


--
-- TOC entry 5636 (class 2606 OID 25013)
-- Name: pedido_pagamento pedpg_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_pagamento
    ADD CONSTRAINT pedpg_id_pkey PRIMARY KEY (pedpg_id);


--
-- TOC entry 5500 (class 2606 OID 50674)
-- Name: pedido_pagamento pedpg_status_check; Type: CHECK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE public.pedido_pagamento
    ADD CONSTRAINT pedpg_status_check CHECK (((pedpg_status)::text = ANY ((ARRAY['C'::character varying, 'X'::character varying])::text[]))) NOT VALID;


--
-- TOC entry 5547 (class 2606 OID 24694)
-- Name: permissao_item peri_chave_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.permissao_item
    ADD CONSTRAINT peri_chave_pkey PRIMARY KEY (peri_chave);


--
-- TOC entry 5543 (class 2606 OID 24677)
-- Name: permissao_modulo perm_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.permissao_modulo
    ADD CONSTRAINT perm_id_pkey PRIMARY KEY (perm_id);


--
-- TOC entry 5545 (class 2606 OID 24686)
-- Name: permissao_submodulo pers_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.permissao_submodulo
    ADD CONSTRAINT pers_id_pkey PRIMARY KEY (pers_id);


--
-- TOC entry 5576 (class 2606 OID 24803)
-- Name: produto pro_codigo_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_codigo_key UNIQUE (pro_codigo);


--
-- TOC entry 5578 (class 2606 OID 24801)
-- Name: produto pro_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_id_pkey PRIMARY KEY (pro_id);


--
-- TOC entry 5580 (class 2606 OID 24830)
-- Name: produto_empresa proe_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa
    ADD CONSTRAINT proe_id_pkey PRIMARY KEY (proe_id);


--
-- TOC entry 5582 (class 2606 OID 24832)
-- Name: produto_empresa proe_id_produto_proe_id_empresa_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa
    ADD CONSTRAINT proe_id_produto_proe_id_empresa_key UNIQUE (proe_id_produto, proe_id_empresa);


--
-- TOC entry 5700 (class 2606 OID 25349)
-- Name: produto_empresa_departamento_fiscal proedf_codigo_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa_departamento_fiscal
    ADD CONSTRAINT proedf_codigo_pkey PRIMARY KEY (proedf_codigo);


--
-- TOC entry 5616 (class 2606 OID 24940)
-- Name: produto_empresa_grade_item proegi_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa_grade_item
    ADD CONSTRAINT proegi_id_pkey PRIMARY KEY (proegi_id);


--
-- TOC entry 5618 (class 2606 OID 24942)
-- Name: produto_empresa_grade_item proegi_id_produto_empresa_proegi_id_grade_item_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa_grade_item
    ADD CONSTRAINT proegi_id_produto_empresa_proegi_id_grade_item_key UNIQUE (proegi_id_produto_empresa, proegi_id_grade_item);


--
-- TOC entry 5671 (class 2606 OID 25243)
-- Name: produto_empresa_grade_item_fornecedor proegif_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa_grade_item_fornecedor
    ADD CONSTRAINT proegif_id_pkey PRIMARY KEY (proegif_id);


--
-- TOC entry 5600 (class 2606 OID 24882)
-- Name: produto_grupo prog_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_grupo
    ADD CONSTRAINT prog_descricao_key UNIQUE (prog_descricao);


--
-- TOC entry 5602 (class 2606 OID 24880)
-- Name: produto_grupo prog_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_grupo
    ADD CONSTRAINT prog_id_pkey PRIMARY KEY (prog_id);


--
-- TOC entry 5596 (class 2606 OID 24870)
-- Name: produto_localizacao prol_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_localizacao
    ADD CONSTRAINT prol_descricao_key UNIQUE (prol_descricao);


--
-- TOC entry 5598 (class 2606 OID 24868)
-- Name: produto_localizacao prol_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_localizacao
    ADD CONSTRAINT prol_id_pkey PRIMARY KEY (prol_id);


--
-- TOC entry 5608 (class 2606 OID 24904)
-- Name: produto_linha proli_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_linha
    ADD CONSTRAINT proli_descricao_key UNIQUE (proli_descricao);


--
-- TOC entry 5610 (class 2606 OID 24902)
-- Name: produto_linha proli_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_linha
    ADD CONSTRAINT proli_id_pkey PRIMARY KEY (proli_id);


--
-- TOC entry 5588 (class 2606 OID 24851)
-- Name: produto_marca prom_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_marca
    ADD CONSTRAINT prom_descricao_key UNIQUE (prom_descricao);


--
-- TOC entry 5590 (class 2606 OID 24849)
-- Name: produto_marca prom_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_marca
    ADD CONSTRAINT prom_id_pkey PRIMARY KEY (prom_id);


--
-- TOC entry 5592 (class 2606 OID 24858)
-- Name: produto_medida prome_codigo_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_medida
    ADD CONSTRAINT prome_codigo_pkey PRIMARY KEY (prome_codigo);


--
-- TOC entry 5594 (class 2606 OID 24860)
-- Name: produto_medida prome_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_medida
    ADD CONSTRAINT prome_descricao_key UNIQUE (prome_descricao);


--
-- TOC entry 5739 (class 2606 OID 59112)
-- Name: produto_opcao proo_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao
    ADD CONSTRAINT proo_descricao_key UNIQUE (proo_descricao);


--
-- TOC entry 5741 (class 2606 OID 59110)
-- Name: produto_opcao proo_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao
    ADD CONSTRAINT proo_id_pkey PRIMARY KEY (proo_id);


--
-- TOC entry 5743 (class 2606 OID 59116)
-- Name: produto_opcao_item prooi_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_item
    ADD CONSTRAINT prooi_id_pkey PRIMARY KEY (prooi_id);


--
-- TOC entry 5746 (class 2606 OID 59118)
-- Name: produto_opcao_item prooi_id_produto_opcao_prooi_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_item
    ADD CONSTRAINT prooi_id_produto_opcao_prooi_descricao_key UNIQUE (prooi_id_produto_opcao, prooi_descricao);


--
-- TOC entry 5748 (class 2606 OID 59135)
-- Name: produto_opcao_produto proop_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_produto
    ADD CONSTRAINT proop_id_pkey PRIMARY KEY (proop_id);


--
-- TOC entry 5751 (class 2606 OID 59137)
-- Name: produto_opcao_produto proop_id_produto_proop_id_produto_opcao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_produto
    ADD CONSTRAINT proop_id_produto_proop_id_produto_opcao_key UNIQUE (proop_id_produto, proop_id_produto_opcao);


--
-- TOC entry 5604 (class 2606 OID 24893)
-- Name: produto_subgrupo pros_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_subgrupo
    ADD CONSTRAINT pros_descricao_key UNIQUE (pros_descricao);


--
-- TOC entry 5606 (class 2606 OID 24891)
-- Name: produto_subgrupo pros_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_subgrupo
    ADD CONSTRAINT pros_id_pkey PRIMARY KEY (pros_id);


--
-- TOC entry 5584 (class 2606 OID 24839)
-- Name: produto_tipo prot_codigo_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_tipo
    ADD CONSTRAINT prot_codigo_pkey PRIMARY KEY (prot_codigo);


--
-- TOC entry 5586 (class 2606 OID 24841)
-- Name: produto_tipo prot_descricao_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_tipo
    ADD CONSTRAINT prot_descricao_key UNIQUE (prot_descricao);


--
-- TOC entry 5523 (class 2606 OID 24584)
-- Name: sistema sis_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.sistema
    ADD CONSTRAINT sis_id_pkey PRIMARY KEY (sis_id);


--
-- TOC entry 5646 (class 2606 OID 25052)
-- Name: situacao_tributaria sitt_chave_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.situacao_tributaria
    ADD CONSTRAINT sitt_chave_key UNIQUE (sitt_codigo_cfop_origem, sitt_codigo_cfop_destino, sitt_codigo_produto_tipo, sitt_cst);


--
-- TOC entry 5648 (class 2606 OID 25050)
-- Name: situacao_tributaria sitt_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.situacao_tributaria
    ADD CONSTRAINT sitt_id_pkey PRIMARY KEY (sitt_id);


--
-- TOC entry 5679 (class 2606 OID 25272)
-- Name: terminal ter_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal
    ADD CONSTRAINT ter_id_pkey PRIMARY KEY (ter_id);


--
-- TOC entry 5681 (class 2606 OID 34266)
-- Name: terminal ter_numero_venda_pendente_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal
    ADD CONSTRAINT ter_numero_venda_pendente_key UNIQUE (ter_numero_venda_pendente);


--
-- TOC entry 5683 (class 2606 OID 25300)
-- Name: terminal_historico terh_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico
    ADD CONSTRAINT terh_id_pkey PRIMARY KEY (terh_id);


--
-- TOC entry 5687 (class 2606 OID 50697)
-- Name: terminal_historico terh_id_terminal_terh_data_terh_sequencia_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico
    ADD CONSTRAINT terh_id_terminal_terh_data_terh_sequencia_key UNIQUE (terh_id_terminal, terh_data, terh_sequencia);


--
-- TOC entry 5735 (class 2606 OID 50654)
-- Name: terminal_historico_conferencia terhc_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico_conferencia
    ADD CONSTRAINT terhc_id_pkey PRIMARY KEY (terhc_id);


--
-- TOC entry 5737 (class 2606 OID 50656)
-- Name: terminal_historico_conferencia terhc_id_terminal_historico_terhc_forma_pagamento_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico_conferencia
    ADD CONSTRAINT terhc_id_terminal_historico_terhc_forma_pagamento_key UNIQUE (terhc_id_terminal_historico, terhc_forma_pagamento);


--
-- TOC entry 5694 (class 2606 OID 25334)
-- Name: uf_produto ufp_id_pkey; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.uf_produto
    ADD CONSTRAINT ufp_id_pkey PRIMARY KEY (ufp_id);


--
-- TOC entry 5696 (class 2606 OID 25336)
-- Name: uf_produto ufp_uf_ufp_codigo_item_ipm_key; Type: CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.uf_produto
    ADD CONSTRAINT ufp_uf_ufp_codigo_item_ipm_key UNIQUE (ufp_uf, ufp_codigo_item_ipm);


--
-- TOC entry 5572 (class 1259 OID 25978)
-- Name: clie_id_cliente_clie_principal_key; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX clie_id_cliente_clie_principal_key ON public.cliente_endereco USING btree (clie_id_cliente) WHERE (clie_principal = true);


--
-- TOC entry 5717 (class 1259 OID 50716)
-- Name: cntpg_id_pagamento_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX cntpg_id_pagamento_idx ON public.conta_pagamento USING btree (cntpg_id_pagamento);


--
-- TOC entry 5690 (class 1259 OID 25974)
-- Name: contc_chave_key; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX contc_chave_key ON public.contrato_cliente USING btree (contc_id_contrato, contc_id_cliente, contc_id_terminal) WHERE (((contc_status)::text = 'A'::text) OR ((contc_status)::text = 'I'::text));


--
-- TOC entry 5619 (class 1259 OID 25976)
-- Name: cred_id_contrato_cliente_ped_data_key; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX cred_id_contrato_cliente_ped_data_key ON public.pedido USING btree (cred_id_contrato_cliente, ped_data) WHERE (((ped_tipo)::text = 'C'::text) AND (cred_id_contrato_cliente IS NOT NULL));


--
-- TOC entry 5755 (class 1259 OID 59203)
-- Name: estm_id_empresa_estm_data_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX estm_id_empresa_estm_data_idx ON public.estoque_movimento USING btree (estm_id_empresa, estm_data);


--
-- TOC entry 5760 (class 1259 OID 59201)
-- Name: estm_nota_fiscal_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX estm_nota_fiscal_idx ON public.estoque_movimento USING btree (estm_id_nota_fiscal) WHERE (estm_id_nota_fiscal IS NOT NULL);


--
-- TOC entry 5761 (class 1259 OID 59202)
-- Name: estm_pedido_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX estm_pedido_idx ON public.estoque_movimento USING btree (estm_numero_pedido) WHERE (estm_numero_pedido IS NOT NULL);


--
-- TOC entry 5766 (class 1259 OID 59248)
-- Name: estmi_id_produto_empresa_grade_item_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX estmi_id_produto_empresa_grade_item_idx ON public.estoque_movimento_item USING btree (estmi_id_produto_empresa_grade_item);


--
-- TOC entry 5767 (class 1259 OID 59246)
-- Name: estmi_nota_fiscal_item_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX estmi_nota_fiscal_item_idx ON public.estoque_movimento_item USING btree (estmi_id_nota_fiscal_item) WHERE (estmi_id_nota_fiscal_item IS NOT NULL);


--
-- TOC entry 5768 (class 1259 OID 59247)
-- Name: estmi_pedido_item_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX estmi_pedido_item_idx ON public.estoque_movimento_item USING btree (estmi_id_pedido_item, estmi_id_produto_empresa_grade_item, estmi_tipo) WHERE (estmi_id_pedido_item IS NOT NULL);


--
-- TOC entry 5659 (class 1259 OID 50713)
-- Name: pag_id_cliente_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX pag_id_cliente_idx ON public.pagamento USING btree (pag_id_cliente);


--
-- TOC entry 5662 (class 1259 OID 50712)
-- Name: pag_id_terminal_historico_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX pag_id_terminal_historico_idx ON public.pagamento USING btree (pag_id_terminal_historico);


--
-- TOC entry 5663 (class 1259 OID 50714)
-- Name: pagi_id_pagamento_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX pagi_id_pagamento_idx ON public.pagamento_item USING btree (pagi_id_pagamento);


--
-- TOC entry 5620 (class 1259 OID 50711)
-- Name: ped_id_terminal_historico_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX ped_id_terminal_historico_idx ON public.pedido USING btree (ped_id_terminal_historico);


--
-- TOC entry 5621 (class 1259 OID 50683)
-- Name: ped_id_terminal_ped_data_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX ped_id_terminal_ped_data_idx ON public.pedido USING btree (ped_id_terminal, ped_data);


--
-- TOC entry 5627 (class 1259 OID 50684)
-- Name: pedi_numero_pedido_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX pedi_numero_pedido_idx ON public.pedido_item USING btree (pedi_numero_pedido);


--
-- TOC entry 5752 (class 1259 OID 59164)
-- Name: pedio_id_pedido_item_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX pedio_id_pedido_item_idx ON public.pedido_item_opcao USING btree (pedio_id_pedido_item);


--
-- TOC entry 5634 (class 1259 OID 50715)
-- Name: pedpg_id_pagamento_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX pedpg_id_pagamento_idx ON public.pedido_pagamento USING btree (pedpg_id_pagamento);


--
-- TOC entry 5637 (class 1259 OID 59002)
-- Name: pedpg_liquidacao_key; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX pedpg_liquidacao_key ON public.pedido_pagamento USING btree (pedpg_numero_pedido, pedpg_id_pagamento, pedpg_valor, pedpg_data) WHERE ((pedpg_status)::text = 'C'::text);


--
-- TOC entry 5638 (class 1259 OID 50734)
-- Name: pedpg_numero_pedido_pedpg_id_pagamento_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX pedpg_numero_pedido_pedpg_id_pagamento_idx ON public.pedido_pagamento USING btree (pedpg_numero_pedido, pedpg_id_pagamento);


--
-- TOC entry 5744 (class 1259 OID 59162)
-- Name: prooi_id_produto_opcao_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX prooi_id_produto_opcao_idx ON public.produto_opcao_item USING btree (prooi_id_produto_opcao);


--
-- TOC entry 5749 (class 1259 OID 59163)
-- Name: proop_id_produto_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX proop_id_produto_idx ON public.produto_opcao_produto USING btree (proop_id_produto);


--
-- TOC entry 5684 (class 1259 OID 50710)
-- Name: terh_id_terminal_terh_data_abertura_idx; Type: INDEX; Schema: public; Owner: postgres
--

CREATE INDEX terh_id_terminal_terh_data_abertura_idx ON public.terminal_historico USING btree (terh_id_terminal, terh_data_abertura DESC);


--
-- TOC entry 5685 (class 1259 OID 50709)
-- Name: terh_id_terminal_terh_data_terh_fechado_key; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX terh_id_terminal_terh_data_terh_fechado_key ON public.terminal_historico USING btree (terh_id_terminal, terh_data) WHERE (terh_fechado IS FALSE);


--
-- TOC entry 5624 (class 1259 OID 34290)
-- Name: ven_id_comanda_ped_status_key; Type: INDEX; Schema: public; Owner: postgres
--

CREATE UNIQUE INDEX ven_id_comanda_ped_status_key ON public.pedido USING btree (ven_id_comanda) WHERE ((ped_status)::text <> ALL ((ARRAY['C'::character varying, 'X'::character varying])::text[]));


--
-- TOC entry 5897 (class 2620 OID 50732)
-- Name: pedido trg_atualizar_cliente_crediario; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_cliente_crediario AFTER UPDATE OF ped_id_cliente ON public.pedido FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_cliente_crediario();


--
-- TOC entry 5920 (class 2620 OID 25989)
-- Name: comanda trg_atualizar_data_abertura_comanda; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_data_abertura_comanda BEFORE UPDATE OF com_status ON public.comanda FOR EACH ROW WHEN (((((old.com_status)::text = 'L'::text) AND ((new.com_status)::text = 'C'::text)) OR (((old.com_status)::text <> 'L'::text) AND ((new.com_status)::text = 'L'::text)))) EXECUTE FUNCTION public.fn_atualizar_data_abertura_comanda();


--
-- TOC entry 5898 (class 2620 OID 59275)
-- Name: pedido trg_atualizar_estoque_movimento_pedido; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_estoque_movimento_pedido AFTER UPDATE OF ped_status, ped_tipo ON public.pedido FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_estoque_movimento_pedido();


--
-- TOC entry 5903 (class 2620 OID 59271)
-- Name: pedido_item trg_atualizar_estoque_movimento_pedido_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_estoque_movimento_pedido_item AFTER INSERT OR UPDATE OF pedi_quantidade, pedi_id_produto_empresa_grade_item, pedi_status ON public.pedido_item FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_estoque_movimento_pedido_item();


--
-- TOC entry 5912 (class 2620 OID 59263)
-- Name: nota_fiscal_item trg_atualizar_quantidade_estoque_nota_fiscal; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_quantidade_estoque_nota_fiscal AFTER INSERT OR UPDATE ON public.nota_fiscal_item FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_quantidade_estoque_nota_fiscal();


--
-- TOC entry 5904 (class 2620 OID 26013)
-- Name: pedido_item trg_atualizar_quantidade_itens_pedido; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_quantidade_itens_pedido AFTER INSERT OR DELETE ON public.pedido_item FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_quantidade_itens_pedido();


--
-- TOC entry 5899 (class 2620 OID 34291)
-- Name: pedido trg_atualizar_status_comanda; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_status_comanda AFTER INSERT OR UPDATE OF ven_id_comanda, ped_status ON public.pedido FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_status_comanda();


--
-- TOC entry 5917 (class 2620 OID 26063)
-- Name: pagamento_item trg_atualizar_totais_pagamento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_totais_pagamento AFTER INSERT OR DELETE OR UPDATE ON public.pagamento_item FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_totais_pagamento();


--
-- TOC entry 5900 (class 2620 OID 59049)
-- Name: pedido trg_atualizar_valor_consumo_comanda; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_valor_consumo_comanda AFTER INSERT OR UPDATE OF ped_valor_total, ped_status ON public.pedido FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_valor_consumo_comanda();


--
-- TOC entry 5914 (class 2620 OID 26007)
-- Name: pagamento trg_atualizar_valor_credito_cliente_por_pagamento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_valor_credito_cliente_por_pagamento AFTER INSERT OR DELETE OR UPDATE ON public.pagamento FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_valor_credito_cliente_por_pagamento();


--
-- TOC entry 5908 (class 2620 OID 25993)
-- Name: crediario_parcela trg_atualizar_valor_limite_disponivel_cliente; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_valor_limite_disponivel_cliente AFTER INSERT OR DELETE OR UPDATE ON public.crediario_parcela FOR EACH ROW EXECUTE FUNCTION public.fn_atualizar_valor_limite_disponivel_cliente();


--
-- TOC entry 5905 (class 2620 OID 26017)
-- Name: pedido_item trg_atualizar_valores_pedido; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_atualizar_valores_pedido AFTER UPDATE OF pedi_status ON public.pedido_item FOR EACH ROW WHEN (((((old.pedi_status)::text <> 'X'::text) AND ((new.pedi_status)::text = 'X'::text)) OR (((old.pedi_status)::text = 'X'::text) AND ((new.pedi_status)::text <> 'X'::text)))) EXECUTE FUNCTION public.fn_atualizar_valores_pedido();


--
-- TOC entry 5921 (class 2620 OID 59254)
-- Name: estoque_movimento trg_bloquear_exclusao_estoque_movimento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_bloquear_exclusao_estoque_movimento BEFORE DELETE ON public.estoque_movimento FOR EACH ROW EXECUTE FUNCTION public.fn_bloquear_exclusao_estoque_movimento();


--
-- TOC entry 5924 (class 2620 OID 59256)
-- Name: estoque_movimento_item trg_bloquear_exclusao_estoque_movimento_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_bloquear_exclusao_estoque_movimento_item BEFORE DELETE ON public.estoque_movimento_item FOR EACH ROW EXECUTE FUNCTION public.fn_bloquear_exclusao_estoque_movimento_item();


--
-- TOC entry 5915 (class 2620 OID 26009)
-- Name: pagamento trg_calcular_totais_pagamento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_calcular_totais_pagamento BEFORE INSERT OR UPDATE ON public.pagamento FOR EACH ROW EXECUTE FUNCTION public.fn_calcular_totais_pagamento();


--
-- TOC entry 5910 (class 2620 OID 59267)
-- Name: nota_fiscal trg_cancelar_estoque_movimento_nota_fiscal; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_cancelar_estoque_movimento_nota_fiscal BEFORE DELETE ON public.nota_fiscal FOR EACH ROW EXECUTE FUNCTION public.fn_cancelar_estoque_movimento_nota_fiscal();


--
-- TOC entry 5901 (class 2620 OID 59277)
-- Name: pedido trg_cancelar_estoque_movimento_pedido; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_cancelar_estoque_movimento_pedido BEFORE DELETE ON public.pedido FOR EACH ROW EXECUTE FUNCTION public.fn_cancelar_estoque_movimento_pedido();


--
-- TOC entry 5890 (class 2620 OID 26071)
-- Name: cliente_fornecedor trg_definir_codigo_cliente_fornecedor; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_definir_codigo_cliente_fornecedor BEFORE INSERT ON public.cliente_fornecedor FOR EACH ROW EXECUTE FUNCTION public.fn_definir_codigo_cliente_fornecedor();


--
-- TOC entry 5892 (class 2620 OID 26029)
-- Name: produto trg_definir_codigo_produto; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_definir_codigo_produto BEFORE INSERT ON public.produto FOR EACH ROW EXECUTE FUNCTION public.fn_definir_codigo_produto();


--
-- TOC entry 5922 (class 2620 OID 59250)
-- Name: estoque_movimento trg_definir_numero_estoque_movimento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_definir_numero_estoque_movimento BEFORE INSERT ON public.estoque_movimento FOR EACH ROW EXECUTE FUNCTION public.fn_definir_numero_estoque_movimento();


--
-- TOC entry 5925 (class 2620 OID 59252)
-- Name: estoque_movimento_item trg_definir_numero_estoque_movimento_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_definir_numero_estoque_movimento_item BEFORE INSERT ON public.estoque_movimento_item FOR EACH ROW EXECUTE FUNCTION public.fn_definir_numero_estoque_movimento_item();


--
-- TOC entry 5891 (class 2620 OID 25987)
-- Name: cliente_fornecedor trg_definir_valor_limite_disponivel_cliente; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_definir_valor_limite_disponivel_cliente BEFORE INSERT OR UPDATE ON public.cliente_fornecedor FOR EACH ROW EXECUTE FUNCTION public.fn_definir_valor_limite_disponivel_cliente();


--
-- TOC entry 5913 (class 2620 OID 59265)
-- Name: nota_fiscal_item trg_excluir_estoque_movimento_nota_fiscal_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_excluir_estoque_movimento_nota_fiscal_item BEFORE DELETE ON public.nota_fiscal_item FOR EACH ROW EXECUTE FUNCTION public.fn_excluir_estoque_movimento_nota_fiscal_item();


--
-- TOC entry 5906 (class 2620 OID 59273)
-- Name: pedido_item trg_excluir_estoque_movimento_pedido_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_excluir_estoque_movimento_pedido_item BEFORE DELETE ON public.pedido_item FOR EACH ROW EXECUTE FUNCTION public.fn_excluir_estoque_movimento_pedido_item();


--
-- TOC entry 5887 (class 2620 OID 25995)
-- Name: empresa trg_inserir_parametro_empresa_apos_empresa; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_inserir_parametro_empresa_apos_empresa AFTER INSERT ON public.empresa FOR EACH ROW EXECUTE FUNCTION public.fn_inserir_parametro_empresa_apos_empresa();


--
-- TOC entry 5889 (class 2620 OID 26011)
-- Name: parametro trg_inserir_parametro_empresa_apos_parametro; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_inserir_parametro_empresa_apos_parametro AFTER INSERT ON public.parametro FOR EACH ROW EXECUTE FUNCTION public.fn_inserir_parametro_empresa_apos_parametro();


--
-- TOC entry 5888 (class 2620 OID 25997)
-- Name: empresa trg_inserir_produto_empresa_apos_empresa; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_inserir_produto_empresa_apos_empresa AFTER INSERT ON public.empresa FOR EACH ROW EXECUTE FUNCTION public.fn_inserir_produto_empresa_apos_empresa();


--
-- TOC entry 5894 (class 2620 OID 25999)
-- Name: grade_item trg_inserir_produto_empresa_grade_item_apos_grade_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_inserir_produto_empresa_grade_item_apos_grade_item AFTER INSERT ON public.grade_item FOR EACH ROW EXECUTE FUNCTION public.fn_inserir_produto_empresa_grade_item_apos_grade_item();


--
-- TOC entry 5893 (class 2620 OID 26027)
-- Name: produto_empresa trg_inserir_produto_empresa_grade_item_apos_produto_empresa; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_inserir_produto_empresa_grade_item_apos_produto_empresa AFTER INSERT ON public.produto_empresa FOR EACH ROW EXECUTE FUNCTION public.fn_inserir_produto_empresa_grade_item_apos_produto_empresa();


--
-- TOC entry 5895 (class 2620 OID 59281)
-- Name: produto_empresa_grade_item trg_lancar_saldo_inicial_produto_empresa_grade_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_lancar_saldo_inicial_produto_empresa_grade_item AFTER INSERT ON public.produto_empresa_grade_item FOR EACH ROW WHEN (((COALESCE(new.proegi_quantidade_estoque, (0)::numeric) <> (0)::numeric) OR (COALESCE(new.proegi_quantidade_prateleira, (0)::numeric) <> (0)::numeric))) EXECUTE FUNCTION public.fn_lancar_saldo_inicial_produto_empresa_grade_item();


--
-- TOC entry 5902 (class 2620 OID 34268)
-- Name: pedido trg_liberar_venda_pendente_terminal; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_liberar_venda_pendente_terminal AFTER UPDATE OF ped_status ON public.pedido FOR EACH ROW EXECUTE FUNCTION public.fn_liberar_venda_pendente_terminal();


--
-- TOC entry 5923 (class 2620 OID 59261)
-- Name: estoque_movimento trg_projetar_estoque_movimento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_projetar_estoque_movimento AFTER UPDATE OF estm_status ON public.estoque_movimento FOR EACH ROW EXECUTE FUNCTION public.fn_projetar_estoque_movimento();


--
-- TOC entry 5926 (class 2620 OID 59259)
-- Name: estoque_movimento_item trg_projetar_estoque_movimento_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_projetar_estoque_movimento_item AFTER INSERT OR DELETE OR UPDATE ON public.estoque_movimento_item FOR EACH ROW EXECUTE FUNCTION public.fn_projetar_estoque_movimento_item();


--
-- TOC entry 5896 (class 2620 OID 59279)
-- Name: produto_empresa_grade_item trg_proteger_saldo_produto_empresa_grade_item; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_proteger_saldo_produto_empresa_grade_item BEFORE UPDATE OF proegi_quantidade_estoque, proegi_quantidade_prateleira ON public.produto_empresa_grade_item FOR EACH ROW EXECUTE FUNCTION public.fn_proteger_saldo_produto_empresa_grade_item();


--
-- TOC entry 5919 (class 2620 OID 50722)
-- Name: conta_pagamento trg_recalcular_credito_pagamento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_recalcular_credito_pagamento AFTER INSERT OR DELETE OR UPDATE ON public.conta_pagamento FOR EACH ROW EXECUTE FUNCTION public.fn_recalcular_credito_pagamento_trigger();


--
-- TOC entry 5918 (class 2620 OID 50720)
-- Name: pagamento_item trg_recalcular_credito_pagamento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_recalcular_credito_pagamento AFTER INSERT OR DELETE OR UPDATE ON public.pagamento_item FOR EACH ROW EXECUTE FUNCTION public.fn_recalcular_credito_pagamento_trigger();


--
-- TOC entry 5909 (class 2620 OID 50721)
-- Name: pedido_pagamento trg_recalcular_credito_pagamento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_recalcular_credito_pagamento AFTER INSERT OR DELETE OR UPDATE ON public.pedido_pagamento FOR EACH ROW EXECUTE FUNCTION public.fn_recalcular_credito_pagamento_trigger();


--
-- TOC entry 5907 (class 2620 OID 26019)
-- Name: pedido_item trg_validar_comanda_bloqueada; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_validar_comanda_bloqueada BEFORE INSERT ON public.pedido_item FOR EACH ROW EXECUTE FUNCTION public.fn_validar_comanda_bloqueada();


--
-- TOC entry 5916 (class 2620 OID 50724)
-- Name: pagamento trg_validar_credito_pagamento; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE CONSTRAINT TRIGGER trg_validar_credito_pagamento AFTER INSERT OR UPDATE ON public.pagamento DEFERRABLE INITIALLY DEFERRED FOR EACH ROW EXECUTE FUNCTION public.fn_validar_credito_pagamento();


--
-- TOC entry 5911 (class 2620 OID 26003)
-- Name: nota_fiscal trg_verificar_cancelamento_nota_fiscal; Type: TRIGGER; Schema: public; Owner: postgres
--

CREATE TRIGGER trg_verificar_cancelamento_nota_fiscal AFTER UPDATE OF not_status ON public.nota_fiscal FOR EACH ROW EXECUTE FUNCTION public.fn_verificar_cancelamento_nota_fiscal();


--
-- TOC entry 5770 (class 2606 OID 25509)
-- Name: acesso ace_id_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.acesso
    ADD CONSTRAINT ace_id_empresa_fkey FOREIGN KEY (ace_id_empresa) REFERENCES public.empresa(emp_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5771 (class 2606 OID 25504)
-- Name: acesso ace_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.acesso
    ADD CONSTRAINT ace_id_funcionario_fkey FOREIGN KEY (ace_id_funcionario) REFERENCES public.funcionario(fun_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5854 (class 2606 OID 25919)
-- Name: boleto bol_id_crediario_parcela_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.boleto
    ADD CONSTRAINT bol_id_crediario_parcela_fkey FOREIGN KEY (bol_id_crediario_parcela) REFERENCES public.crediario_parcela(crep_id);


--
-- TOC entry 5855 (class 2606 OID 25914)
-- Name: boleto bol_id_funcionario_baixa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.boleto
    ADD CONSTRAINT bol_id_funcionario_baixa_fkey FOREIGN KEY (bol_id_funcionario_baixa) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5856 (class 2606 OID 25909)
-- Name: boleto bol_id_funcionario_emissao_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.boleto
    ADD CONSTRAINT bol_id_funcionario_emissao_fkey FOREIGN KEY (bol_id_funcionario_emissao) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5857 (class 2606 OID 25924)
-- Name: boleto_instrucao boli_id_boleto_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.boleto_instrucao
    ADD CONSTRAINT boli_id_boleto_fkey FOREIGN KEY (boli_id_boleto) REFERENCES public.boleto(bol_id);


--
-- TOC entry 5778 (class 2606 OID 59053)
-- Name: cliente_fornecedor cli_id_cliente_grupo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_fornecedor
    ADD CONSTRAINT cli_id_cliente_grupo_fkey FOREIGN KEY (cli_id_cliente_grupo) REFERENCES public.cliente_grupo(clig_id) ON DELETE SET NULL;


--
-- TOC entry 5779 (class 2606 OID 59058)
-- Name: cliente_fornecedor cli_id_cliente_rota_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_fornecedor
    ADD CONSTRAINT cli_id_cliente_rota_fkey FOREIGN KEY (cli_id_cliente_rota) REFERENCES public.cliente_rota(clir_id) ON DELETE SET NULL;


--
-- TOC entry 5783 (class 2606 OID 25569)
-- Name: cliente_endereco clie_id_cliente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_endereco
    ADD CONSTRAINT clie_id_cliente_fkey FOREIGN KEY (clie_id_cliente) REFERENCES public.cliente_fornecedor(clif_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5780 (class 2606 OID 25544)
-- Name: cliente_fornecedor clif_codigo_pais_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_fornecedor
    ADD CONSTRAINT clif_codigo_pais_fkey FOREIGN KEY (clif_codigo_pais) REFERENCES public.pais(pai_codigo) ON UPDATE CASCADE;


--
-- TOC entry 5781 (class 2606 OID 25549)
-- Name: cliente_fornecedor clif_id_funcionario_cadastro_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_fornecedor
    ADD CONSTRAINT clif_id_funcionario_cadastro_fkey FOREIGN KEY (clif_id_funcionario_cadastro) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5782 (class 2606 OID 25564)
-- Name: cliente_referencia_comercial clirc_id_cliente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.cliente_referencia_comercial
    ADD CONSTRAINT clirc_id_cliente_fkey FOREIGN KEY (clirc_id_cliente) REFERENCES public.cliente_fornecedor(clif_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5858 (class 2606 OID 25939)
-- Name: conta cnt_id_cliente_fornecedor_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta
    ADD CONSTRAINT cnt_id_cliente_fornecedor_fkey FOREIGN KEY (cnt_id_cliente_fornecedor) REFERENCES public.cliente_fornecedor(clif_id);


--
-- TOC entry 5859 (class 2606 OID 25929)
-- Name: conta cnt_id_conta_tipo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta
    ADD CONSTRAINT cnt_id_conta_tipo_fkey FOREIGN KEY (cnt_id_conta_tipo) REFERENCES public.conta_tipo(cntt_id);


--
-- TOC entry 5860 (class 2606 OID 25944)
-- Name: conta cnt_id_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta
    ADD CONSTRAINT cnt_id_empresa_fkey FOREIGN KEY (cnt_id_empresa) REFERENCES public.empresa(emp_id);


--
-- TOC entry 5861 (class 2606 OID 25934)
-- Name: conta cnt_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta
    ADD CONSTRAINT cnt_id_funcionario_fkey FOREIGN KEY (cnt_id_funcionario) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5862 (class 2606 OID 25949)
-- Name: conta_parcela cntp_id_conta_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_parcela
    ADD CONSTRAINT cntp_id_conta_fkey FOREIGN KEY (cntp_id_conta) REFERENCES public.conta(cnt_id);


--
-- TOC entry 5863 (class 2606 OID 25954)
-- Name: conta_pagamento cntpg_id_conta_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_pagamento
    ADD CONSTRAINT cntpg_id_conta_fkey FOREIGN KEY (cntpg_id_conta) REFERENCES public.conta(cnt_id);


--
-- TOC entry 5864 (class 2606 OID 25959)
-- Name: conta_pagamento cntpg_id_pagamento_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_pagamento
    ADD CONSTRAINT cntpg_id_pagamento_fkey FOREIGN KEY (cntpg_id_pagamento) REFERENCES public.pagamento(pag_id);


--
-- TOC entry 5865 (class 2606 OID 25964)
-- Name: conta_recorrencia cntr_id_conta_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_recorrencia
    ADD CONSTRAINT cntr_id_conta_fkey FOREIGN KEY (cntr_id_conta) REFERENCES public.conta(cnt_id);


--
-- TOC entry 5866 (class 2606 OID 25969)
-- Name: comanda com_id_cliente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.comanda
    ADD CONSTRAINT com_id_cliente_fkey FOREIGN KEY (com_id_cliente) REFERENCES public.cliente_fornecedor(clif_id);


--
-- TOC entry 5867 (class 2606 OID 34280)
-- Name: comanda com_id_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.comanda
    ADD CONSTRAINT com_id_empresa_fkey FOREIGN KEY (com_id_empresa) REFERENCES public.empresa(emp_id);


--
-- TOC entry 5868 (class 2606 OID 26044)
-- Name: conta_bancaria_retorno conbr_id_conta_bancaria_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_bancaria_retorno
    ADD CONSTRAINT conbr_id_conta_bancaria_fkey FOREIGN KEY (conbr_id_conta_bancaria) REFERENCES public.conta_bancaria(conb_id);


--
-- TOC entry 5869 (class 2606 OID 26049)
-- Name: conta_bancaria_retorno conbr_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.conta_bancaria_retorno
    ADD CONSTRAINT conbr_id_funcionario_fkey FOREIGN KEY (conbr_id_funcionario) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5851 (class 2606 OID 25894)
-- Name: contrato_cliente contc_id_cliente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.contrato_cliente
    ADD CONSTRAINT contc_id_cliente_fkey FOREIGN KEY (contc_id_cliente) REFERENCES public.cliente_fornecedor(clif_id);


--
-- TOC entry 5852 (class 2606 OID 25899)
-- Name: contrato_cliente contc_id_contrato_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.contrato_cliente
    ADD CONSTRAINT contc_id_contrato_fkey FOREIGN KEY (contc_id_contrato) REFERENCES public.contrato(cont_id);


--
-- TOC entry 5853 (class 2606 OID 25904)
-- Name: contrato_cliente contc_id_terminal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.contrato_cliente
    ADD CONSTRAINT contc_id_terminal_fkey FOREIGN KEY (contc_id_terminal) REFERENCES public.terminal(ter_id);


--
-- TOC entry 5803 (class 2606 OID 25694)
-- Name: pedido cred_id_contrato_cliente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT cred_id_contrato_cliente_fkey FOREIGN KEY (cred_id_contrato_cliente) REFERENCES public.contrato_cliente(contc_id);


--
-- TOC entry 5804 (class 2606 OID 25689)
-- Name: pedido cred_numero_crediario_refinanciado_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT cred_numero_crediario_refinanciado_fkey FOREIGN KEY (cred_numero_crediario_refinanciado) REFERENCES public.pedido(ped_numero);


--
-- TOC entry 5816 (class 2606 OID 25729)
-- Name: crediario_parcela crep_numero_crediario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.crediario_parcela
    ADD CONSTRAINT crep_numero_crediario_fkey FOREIGN KEY (crep_numero_crediario) REFERENCES public.pedido(ped_numero);


--
-- TOC entry 5845 (class 2606 OID 25869)
-- Name: classificacao_tributaria_item ctribi_codigo_classificacao_tributaria_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.classificacao_tributaria_item
    ADD CONSTRAINT ctribi_codigo_classificacao_tributaria_fkey FOREIGN KEY (ctribi_codigo_classificacao_tributaria) REFERENCES public.classificacao_tributaria(ctrib_codigo);


--
-- TOC entry 5879 (class 2606 OID 59181)
-- Name: estoque_movimento estm_id_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento
    ADD CONSTRAINT estm_id_empresa_fkey FOREIGN KEY (estm_id_empresa) REFERENCES public.empresa(emp_id);


--
-- TOC entry 5880 (class 2606 OID 59186)
-- Name: estoque_movimento estm_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento
    ADD CONSTRAINT estm_id_funcionario_fkey FOREIGN KEY (estm_id_funcionario) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5881 (class 2606 OID 59191)
-- Name: estoque_movimento estm_id_nota_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento
    ADD CONSTRAINT estm_id_nota_fiscal_fkey FOREIGN KEY (estm_id_nota_fiscal) REFERENCES public.nota_fiscal(not_id) ON DELETE SET NULL;


--
-- TOC entry 5882 (class 2606 OID 59196)
-- Name: estoque_movimento estm_numero_pedido_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento
    ADD CONSTRAINT estm_numero_pedido_fkey FOREIGN KEY (estm_numero_pedido) REFERENCES public.pedido(ped_numero) ON DELETE SET NULL;


--
-- TOC entry 5883 (class 2606 OID 59226)
-- Name: estoque_movimento_item estmi_id_estoque_movimento_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento_item
    ADD CONSTRAINT estmi_id_estoque_movimento_fkey FOREIGN KEY (estmi_id_estoque_movimento) REFERENCES public.estoque_movimento(estm_id);


--
-- TOC entry 5884 (class 2606 OID 59236)
-- Name: estoque_movimento_item estmi_id_nota_fiscal_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento_item
    ADD CONSTRAINT estmi_id_nota_fiscal_item_fkey FOREIGN KEY (estmi_id_nota_fiscal_item) REFERENCES public.nota_fiscal_item(notfi_id) ON DELETE SET NULL;


--
-- TOC entry 5885 (class 2606 OID 59241)
-- Name: estoque_movimento_item estmi_id_pedido_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento_item
    ADD CONSTRAINT estmi_id_pedido_item_fkey FOREIGN KEY (estmi_id_pedido_item) REFERENCES public.pedido_item(pedi_id) ON DELETE SET NULL;


--
-- TOC entry 5886 (class 2606 OID 59231)
-- Name: estoque_movimento_item estmi_id_produto_empresa_grade_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.estoque_movimento_item
    ADD CONSTRAINT estmi_id_produto_empresa_grade_item_fkey FOREIGN KEY (estmi_id_produto_empresa_grade_item) REFERENCES public.produto_empresa_grade_item(proegi_id);


--
-- TOC entry 5774 (class 2606 OID 25529)
-- Name: funcionario_cargo_permissao_item fcpi_chave_permissao_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.funcionario_cargo_permissao_item
    ADD CONSTRAINT fcpi_chave_permissao_item_fkey FOREIGN KEY (fcpi_chave_permissao_item) REFERENCES public.permissao_item(peri_chave) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5775 (class 2606 OID 25524)
-- Name: funcionario_cargo_permissao_item fcpi_id_funcionario_cargo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.funcionario_cargo_permissao_item
    ADD CONSTRAINT fcpi_id_funcionario_cargo_fkey FOREIGN KEY (fcpi_id_funcionario_cargo) REFERENCES public.funcionario_cargo(func_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5769 (class 2606 OID 25499)
-- Name: funcionario fun_id_funcionario_cargo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.funcionario
    ADD CONSTRAINT fun_id_funcionario_cargo_fkey FOREIGN KEY (fun_id_funcionario_cargo) REFERENCES public.funcionario_cargo(func_id);


--
-- TOC entry 5800 (class 2606 OID 25654)
-- Name: grade_item grai_id_grade_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.grade_item
    ADD CONSTRAINT grai_id_grade_fkey FOREIGN KEY (grai_id_grade) REFERENCES public.grade(gra_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5822 (class 2606 OID 25774)
-- Name: nota_fiscal not_codigo_cfop_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal
    ADD CONSTRAINT not_codigo_cfop_fkey FOREIGN KEY (not_codigo_cfop) REFERENCES public.cfop(cfop_codigo);


--
-- TOC entry 5823 (class 2606 OID 25764)
-- Name: nota_fiscal not_id_cliente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal
    ADD CONSTRAINT not_id_cliente_fkey FOREIGN KEY (not_id_cliente) REFERENCES public.cliente_fornecedor(clif_id);


--
-- TOC entry 5824 (class 2606 OID 25769)
-- Name: nota_fiscal not_id_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal
    ADD CONSTRAINT not_id_empresa_fkey FOREIGN KEY (not_id_empresa) REFERENCES public.empresa(emp_id);


--
-- TOC entry 5825 (class 2606 OID 25784)
-- Name: nota_fiscal not_id_fornecedor_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal
    ADD CONSTRAINT not_id_fornecedor_fkey FOREIGN KEY (not_id_fornecedor) REFERENCES public.cliente_fornecedor(clif_id);


--
-- TOC entry 5826 (class 2606 OID 25759)
-- Name: nota_fiscal not_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal
    ADD CONSTRAINT not_id_funcionario_fkey FOREIGN KEY (not_id_funcionario) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5827 (class 2606 OID 25779)
-- Name: nota_fiscal not_id_nota_fiscal_contingencia_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal
    ADD CONSTRAINT not_id_nota_fiscal_contingencia_fkey FOREIGN KEY (not_id_nota_fiscal_contingencia) REFERENCES public.nota_fiscal(not_id);


--
-- TOC entry 5842 (class 2606 OID 25854)
-- Name: nota_fiscal_evento notfe_id_nota_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_evento
    ADD CONSTRAINT notfe_id_nota_fiscal_fkey FOREIGN KEY (notfe_id_nota_fiscal) REFERENCES public.nota_fiscal(not_id) ON DELETE CASCADE;


--
-- TOC entry 5830 (class 2606 OID 25809)
-- Name: nota_fiscal_item notfi_codigo_cfop_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_item
    ADD CONSTRAINT notfi_codigo_cfop_fkey FOREIGN KEY (notfi_codigo_cfop) REFERENCES public.cfop(cfop_codigo);


--
-- TOC entry 5831 (class 2606 OID 25799)
-- Name: nota_fiscal_item notfi_id_nota_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_item
    ADD CONSTRAINT notfi_id_nota_fiscal_fkey FOREIGN KEY (notfi_id_nota_fiscal) REFERENCES public.nota_fiscal(not_id) ON DELETE CASCADE;


--
-- TOC entry 5832 (class 2606 OID 25804)
-- Name: nota_fiscal_item notfi_id_produto_empresa_grade_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_item
    ADD CONSTRAINT notfi_id_produto_empresa_grade_item_fkey FOREIGN KEY (notfi_id_produto_empresa_grade_item) REFERENCES public.produto_empresa_grade_item(proegi_id);


--
-- TOC entry 5828 (class 2606 OID 25789)
-- Name: nota_fiscal_parcela notfp_id_nota_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_parcela
    ADD CONSTRAINT notfp_id_nota_fiscal_fkey FOREIGN KEY (notfp_id_nota_fiscal) REFERENCES public.nota_fiscal(not_id) ON DELETE CASCADE;


--
-- TOC entry 5829 (class 2606 OID 25794)
-- Name: nota_fiscal_pagamento notfpa_id_nota_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_pagamento
    ADD CONSTRAINT notfpa_id_nota_fiscal_fkey FOREIGN KEY (notfpa_id_nota_fiscal) REFERENCES public.nota_fiscal(not_id) ON DELETE CASCADE;


--
-- TOC entry 5840 (class 2606 OID 25844)
-- Name: nota_fiscal_referencia notfr_id_nota_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_referencia
    ADD CONSTRAINT notfr_id_nota_fiscal_fkey FOREIGN KEY (notfr_id_nota_fiscal) REFERENCES public.nota_fiscal(not_id) ON DELETE CASCADE;


--
-- TOC entry 5841 (class 2606 OID 25849)
-- Name: nota_fiscal_referencia notfr_id_nota_fiscal_referencia_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_referencia
    ADD CONSTRAINT notfr_id_nota_fiscal_referencia_fkey FOREIGN KEY (notfr_id_nota_fiscal_referencia) REFERENCES public.nota_fiscal(not_id) ON DELETE CASCADE;


--
-- TOC entry 5833 (class 2606 OID 25814)
-- Name: nota_fiscal_transporte_volume notft_id_nota_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.nota_fiscal_transporte_volume
    ADD CONSTRAINT notft_id_nota_fiscal_fkey FOREIGN KEY (notft_id_nota_fiscal) REFERENCES public.nota_fiscal(not_id) ON DELETE CASCADE;


--
-- TOC entry 5834 (class 2606 OID 25824)
-- Name: pagamento pag_id_cliente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pagamento
    ADD CONSTRAINT pag_id_cliente_fkey FOREIGN KEY (pag_id_cliente) REFERENCES public.cliente_fornecedor(clif_id);


--
-- TOC entry 5835 (class 2606 OID 25819)
-- Name: pagamento pag_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pagamento
    ADD CONSTRAINT pag_id_funcionario_fkey FOREIGN KEY (pag_id_funcionario) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5836 (class 2606 OID 25829)
-- Name: pagamento pag_id_terminal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pagamento
    ADD CONSTRAINT pag_id_terminal_fkey FOREIGN KEY (pag_id_terminal) REFERENCES public.terminal(ter_id);


--
-- TOC entry 5837 (class 2606 OID 50704)
-- Name: pagamento pag_id_terminal_historico_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pagamento
    ADD CONSTRAINT pag_id_terminal_historico_fkey FOREIGN KEY (pag_id_terminal_historico) REFERENCES public.terminal_historico(terh_id);


--
-- TOC entry 5838 (class 2606 OID 25839)
-- Name: pagamento_item pagi_id_conta_bancaria_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pagamento_item
    ADD CONSTRAINT pagi_id_conta_bancaria_fkey FOREIGN KEY (pagi_id_conta_bancaria) REFERENCES public.conta_bancaria(conb_id);


--
-- TOC entry 5839 (class 2606 OID 25834)
-- Name: pagamento_item pagi_id_pagamento_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pagamento_item
    ADD CONSTRAINT pagi_id_pagamento_fkey FOREIGN KEY (pagi_id_pagamento) REFERENCES public.pagamento(pag_id);


--
-- TOC entry 5776 (class 2606 OID 25539)
-- Name: parametro_empresa pare_chave_parametro_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.parametro_empresa
    ADD CONSTRAINT pare_chave_parametro_fkey FOREIGN KEY (pare_chave_parametro) REFERENCES public.parametro(par_chave) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5777 (class 2606 OID 25534)
-- Name: parametro_empresa pare_id_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.parametro_empresa
    ADD CONSTRAINT pare_id_empresa_fkey FOREIGN KEY (pare_id_empresa) REFERENCES public.empresa(emp_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5805 (class 2606 OID 25674)
-- Name: pedido ped_id_cliente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT ped_id_cliente_fkey FOREIGN KEY (ped_id_cliente) REFERENCES public.cliente_fornecedor(clif_id);


--
-- TOC entry 5806 (class 2606 OID 25684)
-- Name: pedido ped_id_funcionario_cancelamento_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT ped_id_funcionario_cancelamento_fkey FOREIGN KEY (ped_id_funcionario_cancelamento) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5807 (class 2606 OID 25679)
-- Name: pedido ped_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT ped_id_funcionario_fkey FOREIGN KEY (ped_id_funcionario) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5808 (class 2606 OID 25699)
-- Name: pedido ped_id_nota_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT ped_id_nota_fiscal_fkey FOREIGN KEY (ped_id_nota_fiscal) REFERENCES public.nota_fiscal(not_id);


--
-- TOC entry 5809 (class 2606 OID 25669)
-- Name: pedido ped_id_terminal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT ped_id_terminal_fkey FOREIGN KEY (ped_id_terminal) REFERENCES public.terminal(ter_id);


--
-- TOC entry 5810 (class 2606 OID 50699)
-- Name: pedido ped_id_terminal_historico_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT ped_id_terminal_historico_fkey FOREIGN KEY (ped_id_terminal_historico) REFERENCES public.terminal_historico(terh_id);


--
-- TOC entry 5812 (class 2606 OID 25724)
-- Name: pedido_item pedi_id_funcionario_cancelamento_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item
    ADD CONSTRAINT pedi_id_funcionario_cancelamento_fkey FOREIGN KEY (pedi_id_funcionario_cancelamento) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5813 (class 2606 OID 25719)
-- Name: pedido_item pedi_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item
    ADD CONSTRAINT pedi_id_funcionario_fkey FOREIGN KEY (pedi_id_funcionario) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5814 (class 2606 OID 25714)
-- Name: pedido_item pedi_id_produto_empresa_grade_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item
    ADD CONSTRAINT pedi_id_produto_empresa_grade_item_fkey FOREIGN KEY (pedi_id_produto_empresa_grade_item) REFERENCES public.produto_empresa_grade_item(proegi_id);


--
-- TOC entry 5815 (class 2606 OID 25709)
-- Name: pedido_item pedi_numero_pedido_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item
    ADD CONSTRAINT pedi_numero_pedido_fkey FOREIGN KEY (pedi_numero_pedido) REFERENCES public.pedido(ped_numero);


--
-- TOC entry 5877 (class 2606 OID 59152)
-- Name: pedido_item_opcao pedio_id_pedido_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item_opcao
    ADD CONSTRAINT pedio_id_pedido_item_fkey FOREIGN KEY (pedio_id_pedido_item) REFERENCES public.pedido_item(pedi_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5878 (class 2606 OID 59157)
-- Name: pedido_item_opcao pedio_id_produto_empresa_grade_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_item_opcao
    ADD CONSTRAINT pedio_id_produto_empresa_grade_item_fkey FOREIGN KEY (pedio_id_produto_empresa_grade_item) REFERENCES public.produto_empresa_grade_item(proegi_id) ON UPDATE SET NULL ON DELETE SET NULL;


--
-- TOC entry 5817 (class 2606 OID 25739)
-- Name: pedido_pagamento pedpg_id_pagamento_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_pagamento
    ADD CONSTRAINT pedpg_id_pagamento_fkey FOREIGN KEY (pedpg_id_pagamento) REFERENCES public.pagamento(pag_id);


--
-- TOC entry 5818 (class 2606 OID 25734)
-- Name: pedido_pagamento pedpg_numero_pedido_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido_pagamento
    ADD CONSTRAINT pedpg_numero_pedido_fkey FOREIGN KEY (pedpg_numero_pedido) REFERENCES public.pedido(ped_numero);


--
-- TOC entry 5773 (class 2606 OID 25519)
-- Name: permissao_item peri_id_permissao_submodulo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.permissao_item
    ADD CONSTRAINT peri_id_permissao_submodulo_fkey FOREIGN KEY (peri_id_permissao_submodulo) REFERENCES public.permissao_submodulo(pers_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5772 (class 2606 OID 25514)
-- Name: permissao_submodulo pers_id_permissao_modulo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.permissao_submodulo
    ADD CONSTRAINT pers_id_permissao_modulo_fkey FOREIGN KEY (pers_id_permissao_modulo) REFERENCES public.permissao_modulo(perm_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5784 (class 2606 OID 25614)
-- Name: produto pro_codigo_ncm_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_codigo_ncm_fkey FOREIGN KEY (pro_codigo_ncm) REFERENCES public.ncm(ncm_codigo);


--
-- TOC entry 5785 (class 2606 OID 25574)
-- Name: produto pro_codigo_produto_medida_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_codigo_produto_medida_fkey FOREIGN KEY (pro_codigo_produto_medida) REFERENCES public.produto_medida(prome_codigo) ON UPDATE CASCADE;


--
-- TOC entry 5786 (class 2606 OID 25584)
-- Name: produto pro_codigo_produto_tipo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_codigo_produto_tipo_fkey FOREIGN KEY (pro_codigo_produto_tipo) REFERENCES public.produto_tipo(prot_codigo);


--
-- TOC entry 5787 (class 2606 OID 25619)
-- Name: produto pro_id_classificacao_tributaria_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_id_classificacao_tributaria_item_fkey FOREIGN KEY (pro_id_classificacao_tributaria_item) REFERENCES public.classificacao_tributaria_item(ctribi_id);


--
-- TOC entry 5788 (class 2606 OID 25609)
-- Name: produto pro_id_grade_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_id_grade_fkey FOREIGN KEY (pro_id_grade) REFERENCES public.grade(gra_id);


--
-- TOC entry 5789 (class 2606 OID 25594)
-- Name: produto pro_id_produto_grupo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_id_produto_grupo_fkey FOREIGN KEY (pro_id_produto_grupo) REFERENCES public.produto_grupo(prog_id) ON UPDATE SET NULL ON DELETE SET NULL;


--
-- TOC entry 5790 (class 2606 OID 25599)
-- Name: produto pro_id_produto_linha_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_id_produto_linha_fkey FOREIGN KEY (pro_id_produto_linha) REFERENCES public.produto_linha(proli_id) ON UPDATE SET NULL ON DELETE SET NULL;


--
-- TOC entry 5791 (class 2606 OID 25589)
-- Name: produto pro_id_produto_localizacao_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_id_produto_localizacao_fkey FOREIGN KEY (pro_id_produto_localizacao) REFERENCES public.produto_localizacao(prol_id) ON UPDATE SET NULL ON DELETE SET NULL;


--
-- TOC entry 5792 (class 2606 OID 25579)
-- Name: produto pro_id_produto_marca_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_id_produto_marca_fkey FOREIGN KEY (pro_id_produto_marca) REFERENCES public.produto_marca(prom_id) ON UPDATE SET NULL ON DELETE SET NULL;


--
-- TOC entry 5793 (class 2606 OID 25604)
-- Name: produto pro_id_produto_subgrupo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto
    ADD CONSTRAINT pro_id_produto_subgrupo_fkey FOREIGN KEY (pro_id_produto_subgrupo) REFERENCES public.produto_subgrupo(pros_id) ON UPDATE SET NULL ON DELETE SET NULL;


--
-- TOC entry 5794 (class 2606 OID 25639)
-- Name: produto_empresa proe_codigo_produto_empresa_departamento_fiscal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa
    ADD CONSTRAINT proe_codigo_produto_empresa_departamento_fiscal_fkey FOREIGN KEY (proe_codigo_produto_empresa_departamento_fiscal) REFERENCES public.produto_empresa_departamento_fiscal(proedf_codigo);


--
-- TOC entry 5795 (class 2606 OID 25629)
-- Name: produto_empresa proe_id_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa
    ADD CONSTRAINT proe_id_empresa_fkey FOREIGN KEY (proe_id_empresa) REFERENCES public.empresa(emp_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5796 (class 2606 OID 25624)
-- Name: produto_empresa proe_id_produto_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa
    ADD CONSTRAINT proe_id_produto_fkey FOREIGN KEY (proe_id_produto) REFERENCES public.produto(pro_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5797 (class 2606 OID 25634)
-- Name: produto_empresa proe_id_uf_produto_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa
    ADD CONSTRAINT proe_id_uf_produto_fkey FOREIGN KEY (proe_id_uf_produto) REFERENCES public.uf_produto(ufp_id);


--
-- TOC entry 5801 (class 2606 OID 25664)
-- Name: produto_empresa_grade_item proegi_id_grade_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa_grade_item
    ADD CONSTRAINT proegi_id_grade_item_fkey FOREIGN KEY (proegi_id_grade_item) REFERENCES public.grade_item(grai_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5802 (class 2606 OID 25659)
-- Name: produto_empresa_grade_item proegi_id_produto_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa_grade_item
    ADD CONSTRAINT proegi_id_produto_empresa_fkey FOREIGN KEY (proegi_id_produto_empresa) REFERENCES public.produto_empresa(proe_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5843 (class 2606 OID 25859)
-- Name: produto_empresa_grade_item_fornecedor proegif_id_fornecedor_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa_grade_item_fornecedor
    ADD CONSTRAINT proegif_id_fornecedor_fkey FOREIGN KEY (proegif_id_fornecedor) REFERENCES public.cliente_fornecedor(clif_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5844 (class 2606 OID 25864)
-- Name: produto_empresa_grade_item_fornecedor proegif_id_produto_empresa_grade_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_empresa_grade_item_fornecedor
    ADD CONSTRAINT proegif_id_produto_empresa_grade_item_fkey FOREIGN KEY (proegif_id_produto_empresa_grade_item) REFERENCES public.produto_empresa_grade_item(proegi_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5799 (class 2606 OID 25649)
-- Name: produto_linha proli_id_subgrupo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_linha
    ADD CONSTRAINT proli_id_subgrupo_fkey FOREIGN KEY (proli_id_subgrupo) REFERENCES public.produto_subgrupo(pros_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5872 (class 2606 OID 59129)
-- Name: produto_opcao_item prooi_id_grade_item_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_item
    ADD CONSTRAINT prooi_id_grade_item_fkey FOREIGN KEY (prooi_id_grade_item) REFERENCES public.grade_item(grai_id) ON UPDATE SET NULL ON DELETE SET NULL;


--
-- TOC entry 5873 (class 2606 OID 59124)
-- Name: produto_opcao_item prooi_id_produto_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_item
    ADD CONSTRAINT prooi_id_produto_fkey FOREIGN KEY (prooi_id_produto) REFERENCES public.produto(pro_id) ON UPDATE SET NULL ON DELETE SET NULL;


--
-- TOC entry 5874 (class 2606 OID 59119)
-- Name: produto_opcao_item prooi_id_produto_opcao_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_item
    ADD CONSTRAINT prooi_id_produto_opcao_fkey FOREIGN KEY (prooi_id_produto_opcao) REFERENCES public.produto_opcao(proo_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5875 (class 2606 OID 59138)
-- Name: produto_opcao_produto proop_id_produto_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_produto
    ADD CONSTRAINT proop_id_produto_fkey FOREIGN KEY (proop_id_produto) REFERENCES public.produto(pro_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5876 (class 2606 OID 59143)
-- Name: produto_opcao_produto proop_id_produto_opcao_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_opcao_produto
    ADD CONSTRAINT proop_id_produto_opcao_fkey FOREIGN KEY (proop_id_produto_opcao) REFERENCES public.produto_opcao(proo_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5798 (class 2606 OID 25644)
-- Name: produto_subgrupo pros_id_grupo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.produto_subgrupo
    ADD CONSTRAINT pros_id_grupo_fkey FOREIGN KEY (pros_id_grupo) REFERENCES public.produto_grupo(prog_id) ON UPDATE CASCADE ON DELETE CASCADE;


--
-- TOC entry 5819 (class 2606 OID 25749)
-- Name: situacao_tributaria sitt_codigo_cfop_destino_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.situacao_tributaria
    ADD CONSTRAINT sitt_codigo_cfop_destino_fkey FOREIGN KEY (sitt_codigo_cfop_destino) REFERENCES public.cfop(cfop_codigo);


--
-- TOC entry 5820 (class 2606 OID 25744)
-- Name: situacao_tributaria sitt_codigo_cfop_origem_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.situacao_tributaria
    ADD CONSTRAINT sitt_codigo_cfop_origem_fkey FOREIGN KEY (sitt_codigo_cfop_origem) REFERENCES public.cfop(cfop_codigo);


--
-- TOC entry 5821 (class 2606 OID 25754)
-- Name: situacao_tributaria sitt_codigo_produto_tipo_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.situacao_tributaria
    ADD CONSTRAINT sitt_codigo_produto_tipo_fkey FOREIGN KEY (sitt_codigo_produto_tipo) REFERENCES public.produto_tipo(prot_codigo);


--
-- TOC entry 5846 (class 2606 OID 25874)
-- Name: terminal ter_id_empresa_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal
    ADD CONSTRAINT ter_id_empresa_fkey FOREIGN KEY (ter_id_empresa) REFERENCES public.empresa(emp_id);


--
-- TOC entry 5847 (class 2606 OID 34260)
-- Name: terminal ter_numero_venda_pendente_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal
    ADD CONSTRAINT ter_numero_venda_pendente_fkey FOREIGN KEY (ter_numero_venda_pendente) REFERENCES public.pedido(ped_numero) ON DELETE SET NULL;


--
-- TOC entry 5848 (class 2606 OID 25884)
-- Name: terminal_historico terh_id_funcionario_abertura_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico
    ADD CONSTRAINT terh_id_funcionario_abertura_fkey FOREIGN KEY (terh_id_funcionario_abertura) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5849 (class 2606 OID 25889)
-- Name: terminal_historico terh_id_funcionario_fechamento_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico
    ADD CONSTRAINT terh_id_funcionario_fechamento_fkey FOREIGN KEY (terh_id_funcionario_fechamento) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5850 (class 2606 OID 25879)
-- Name: terminal_historico terh_id_terminal_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico
    ADD CONSTRAINT terh_id_terminal_fkey FOREIGN KEY (terh_id_terminal) REFERENCES public.terminal(ter_id);


--
-- TOC entry 5870 (class 2606 OID 50726)
-- Name: terminal_historico_conferencia terhc_id_funcionario_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico_conferencia
    ADD CONSTRAINT terhc_id_funcionario_fkey FOREIGN KEY (terhc_id_funcionario) REFERENCES public.funcionario(fun_id);


--
-- TOC entry 5871 (class 2606 OID 50657)
-- Name: terminal_historico_conferencia terhc_id_terminal_historico_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.terminal_historico_conferencia
    ADD CONSTRAINT terhc_id_terminal_historico_fkey FOREIGN KEY (terhc_id_terminal_historico) REFERENCES public.terminal_historico(terh_id) ON DELETE CASCADE;


--
-- TOC entry 5811 (class 2606 OID 34285)
-- Name: pedido ven_id_comanda_fkey; Type: FK CONSTRAINT; Schema: public; Owner: postgres
--

ALTER TABLE ONLY public.pedido
    ADD CONSTRAINT ven_id_comanda_fkey FOREIGN KEY (ven_id_comanda) REFERENCES public.comanda(com_id);


-- Completed on 2026-10-05 16:38:31

--
-- PostgreSQL database dump complete
--

\unrestrict NRpncU9th5ziICRTAEM5Egsh6ve0BR28sbt8edqR8JpZ1imXMsm6N82WlxUOaYx

