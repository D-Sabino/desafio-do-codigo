PGDMP     
        	            {            DB_DesafioDoCodigo    15.2    15.2     ?           0    0    ENCODING    ENCODING        SET client_encoding = 'UTF8';
                      false            ?           0    0 
   STDSTRINGS 
   STDSTRINGS     (   SET standard_conforming_strings = 'on';
                      false            ?           0    0 
   SEARCHPATH 
   SEARCHPATH     8   SELECT pg_catalog.set_config('search_path', '', false);
                      false            ?           1262    16548    DB_DesafioDoCodigo    DATABASE     ?   CREATE DATABASE "DB_DesafioDoCodigo" WITH TEMPLATE = template0 ENCODING = 'UTF8' LOCALE_PROVIDER = libc LOCALE = 'Portuguese_Brazil.1252';
 $   DROP DATABASE "DB_DesafioDoCodigo";
                postgres    false            ?            1259    16578    cad_clientes    TABLE     m  CREATE TABLE public.cad_clientes (
    id integer NOT NULL,
    nome character varying(100) NOT NULL,
    endereco character varying(150) NOT NULL,
    cidade character varying(75) NOT NULL,
    estado character varying(50) NOT NULL,
    pais character varying(50) NOT NULL,
    telefone character varying(25) NOT NULL,
    email character varying(100) NOT NULL
);
     DROP TABLE public.cad_clientes;
       public         heap    postgres    false            ?            1259    16577    cad_clientes_id_seq    SEQUENCE     ?   CREATE SEQUENCE public.cad_clientes_id_seq
    AS integer
    START WITH 1
    INCREMENT BY 1
    NO MINVALUE
    NO MAXVALUE
    CACHE 1;
 *   DROP SEQUENCE public.cad_clientes_id_seq;
       public          postgres    false    215            ?           0    0    cad_clientes_id_seq    SEQUENCE OWNED BY     K   ALTER SEQUENCE public.cad_clientes_id_seq OWNED BY public.cad_clientes.id;
          public          postgres    false    214            e           2604    16581    cad_clientes id    DEFAULT     r   ALTER TABLE ONLY public.cad_clientes ALTER COLUMN id SET DEFAULT nextval('public.cad_clientes_id_seq'::regclass);
 >   ALTER TABLE public.cad_clientes ALTER COLUMN id DROP DEFAULT;
       public          postgres    false    214    215    215            ?          0    16578    cad_clientes 
   TABLE DATA           a   COPY public.cad_clientes (id, nome, endereco, cidade, estado, pais, telefone, email) FROM stdin;
    public          postgres    false    215          ?           0    0    cad_clientes_id_seq    SEQUENCE SET     B   SELECT pg_catalog.setval('public.cad_clientes_id_seq', 22, true);
          public          postgres    false    214            g           2606    16585    cad_clientes cad_clientes_pkey 
   CONSTRAINT     \   ALTER TABLE ONLY public.cad_clientes
    ADD CONSTRAINT cad_clientes_pkey PRIMARY KEY (id);
 H   ALTER TABLE ONLY public.cad_clientes DROP CONSTRAINT cad_clientes_pkey;
       public            postgres    false    215            ?   ?  x?m?=r?0???)??ԯ?P?I??"?̏?I?!<?E>NR?H?#?b???C?~?+0?>?.?4??J???0X???QJ?6????T)B?%?c?;C???y ??[s???*µV'?,Ap???o?)??ĄJwl K:?/p?????i???Ǯ̪???)??ײ:????`!??kcՃ)?VqG	??<6Wn%m??'?????m???[???U????NYPI[N?ASHea?$M?Т?w7?ɭ?IH?:?\?????=G*N?Q3Hj??%??TE?[?_Cd쑴t????<6?k????9??[?zO?~? ??:?A?u????c͡/}	1m????،??~??i\?ro%dy???Y?㯼??H?g??r{&5?]???:?T?`yl`??????m
ݚ??w?i?:T6?]!????x???`?f???#????g3p?h	?Q͏?U??*??h??W??>?E ?     