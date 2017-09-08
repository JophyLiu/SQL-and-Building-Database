/*==============================================================*/
/* DBMS name:      Microsoft SQL Server 2012                    */
/* Created on:     2015/12/30 21:37:40                          */
/*==============================================================*/


if exists (select 1
   from sys.sysreferences r join sys.sysobjects o on (o.id = r.constid and o.type = 'F')
   where r.fkeyid = object_id('Staff_attend_info') and o.name = 'FK_STAFF_AT_BELONG2_STAFF_BA')
alter table Staff_attend_info
   drop constraint FK_STAFF_AT_BELONG2_STAFF_BA
go

if exists (select 1
   from sys.sysreferences r join sys.sysobjects o on (o.id = r.constid and o.type = 'F')
   where r.fkeyid = object_id('Staff_mobilize_info') and o.name = 'FK_STAFF_MO_BELONG1_STAFF_BA')
alter table Staff_mobilize_info
   drop constraint FK_STAFF_MO_BELONG1_STAFF_BA
go

if exists (select 1
            from  sysobjects
           where  id = object_id('Staff_attend_info')
            and   type = 'U')
   drop table Staff_attend_info
go

if exists (select 1
            from  sysobjects
           where  id = object_id('Staff_basic_info')
            and   type = 'U')
   drop table Staff_basic_info
go

if exists (select 1
            from  sysindexes
           where  id    = object_id('Staff_mobilize_info')
            and   name  = 'belong1_FK'
            and   indid > 0
            and   indid < 255)
   drop index Staff_mobilize_info.belong1_FK
go

if exists (select 1
            from  sysobjects
           where  id = object_id('Staff_mobilize_info')
            and   type = 'U')
   drop table Staff_mobilize_info
go

if exists (select 1
            from  sysobjects
           where  id = object_id('login')
            and   type = 'U')
   drop table login
go

/*==============================================================*/
/* Table: Staff_attend_info                                     */
/*==============================================================*/
create table Staff_attend_info (
   staff_number         char(7)              not null,
   go_work_time         varchar(8)           not null,
   out_work_time        varchar(8)           not null,
   late_times           int                  null,
   leave_early_times    int                  null,
   in_out               bit                  null,
   sicks                int                  null,
   affair               int                  null,
   leaves_start         datetime             null,
   work_overtime        int                  null,
   overtime_date        datetime             null,
   business_trip        int                  null,
   B_trip_start         datetime             null,
   constraint PK_STAFF_ATTEND_INFO primary key nonclustered (staff_number)
)
go

/*==============================================================*/
/* Table: Staff_basic_info                                      */
/*==============================================================*/
create table Staff_basic_info (
   staff_number         char(7)              not null,
   staff_name           varchar(8)           not null,
   staff_sex            varchar(2)           not null,
   staff_where          varchar(10)          null,
   staff_age            int                  not null,
   staff_birth          datetime             not null,
   staff_add            varchar(20)          null,
   staff_Email          varchar(25)          not null,
   staff_ROFS           varchar(4)           not null,
   staff_major          varchar(12)          not null,
   staff_intime         datetime             not null,
   constraint PK_STAFF_BASIC_INFO primary key nonclustered (staff_number)
)
go

/*==============================================================*/
/* Table: Staff_mobilize_info                                   */
/*==============================================================*/
create table Staff_mobilize_info (
   staff_number         char(7)              not null,
   old_department       varchar(8)           not null,
   new_department       varchar(8)           not null,
   old_position         varchar(10)          not null,
   new_position         varchar(10)          not null,
   out_date             datetime             not null,
   in_date              datetime             not null,
   info                 text                 null,
   constraint PK_STAFF_MOBILIZE_INFO primary key nonclustered (staff_number, old_department, new_department)
)
go

/*==============================================================*/
/* Index: belong1_FK                                            */
/*==============================================================*/
create index belong1_FK on Staff_mobilize_info (
staff_number ASC
)
go

/*==============================================================*/
/* Table: login                                                 */
/*==============================================================*/
create table login (
   "user"               varchar(7)           not null,
   passkey              varchar(6)           not null,
   constraint PK_LOGIN primary key nonclustered ("user", passkey)
)
go

alter table Staff_attend_info
   add constraint FK_STAFF_AT_BELONG2_STAFF_BA foreign key (staff_number)
      references Staff_basic_info (staff_number)
go

alter table Staff_mobilize_info
   add constraint FK_STAFF_MO_BELONG1_STAFF_BA foreign key (staff_number)
      references Staff_basic_info (staff_number)
go

