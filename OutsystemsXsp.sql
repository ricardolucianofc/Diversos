OutSystems X SP
/* 
%LogicalDatabase%=GetLogicalDatabase({PlanoCarreiraProfessores}) */
--Select {Calendario}.* 
--from {Calendario}
EXEC [PlanoCarreiraProfessores].[dbo].[TRIPLO] @VALOR
