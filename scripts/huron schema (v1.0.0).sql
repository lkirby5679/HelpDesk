਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䄀猀猀椀最渀洀攀渀琀猀崀 ⠀ഀഀ
	[AssignmentPK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀䌀愀猀攀吀礀瀀攀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[CatFK] [int] NULL ,਍ऀ嬀刀攀瀀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[IsActive] [bit] NOT NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblCaseTypes] (਍ऀ嬀䌀愀猀攀吀礀瀀攀倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[CaseTypeName] [nvarchar] (32) ,਍ऀ嬀䌀愀猀攀吀礀瀀攀䐀攀猀挀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㄀㈀㠀⤀ Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[RepGroupFK] [int] NULL ,਍ऀ嬀䤀猀䄀挀琀椀瘀攀崀 嬀戀椀琀崀 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[CaseTypeOrder] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblCases] (਍ऀ嬀䌀愀猀攀倀䬀崀 嬀戀椀最椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[ContactFK] [int] NULL ,਍ऀ嬀刀攀瀀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[GroupFK] [int] NULL ,਍ऀ嬀匀琀愀琀甀猀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[CatFK] [int] NULL ,਍ऀ嬀倀爀椀漀爀椀琀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[CaseTypeFK] [int] NULL ,਍ऀ嬀吀椀琀氀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㄀㈀㠀⤀ Ⰰഀഀ
	[Description] [ntext] ,਍ऀ嬀刀攀猀漀氀甀琀椀漀渀崀 嬀渀琀攀砀琀崀 Ⰰഀഀ
	[AltEMail] [nvarchar] (64) ,਍ऀ嬀刀愀椀猀攀搀䐀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[ClosedDate] [datetime] NULL ,਍ऀ嬀䤀猀䄀挀琀椀瘀攀崀 嬀戀椀琀崀 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[EnteredByFK] [int] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ,਍ऀ嬀䌀挀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㘀㐀⤀ Ⰰഀഀ
	[DeptFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀 吀䔀堀吀䤀䴀䄀䜀䔀开伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblCategories] (਍ऀ嬀䌀愀琀倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[CatName] [nvarchar] (32) ,਍ऀ嬀䌀愀琀䐀攀猀挀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㘀㐀⤀ Ⰰഀഀ
	[CatOrder] [int] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ,਍ऀ嬀䤀猀䄀挀琀椀瘀攀崀 嬀戀椀琀崀 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[CaseTypeFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblContacts] (਍ऀ嬀䌀漀渀琀愀挀琀倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[OrgFK] [int] NULL ,਍ऀ嬀䘀一愀洀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[LName] [nvarchar] (32) ,਍ऀ嬀䐀攀瀀琀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[ContactTypeFK] [int] NULL ,਍ऀ嬀䰀愀渀最䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[IsActive] [bit] NOT NULL ,਍ऀ嬀唀猀攀爀一愀洀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㄀㘀⤀ Ⰰഀഀ
	[PW] [nvarchar] (16) ,਍ऀ嬀䤀伀匀琀愀琀甀猀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[IOStatusDate] [datetime] NULL ,਍ऀ嬀䤀伀匀琀愀琀甀猀吀攀砀琀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㈀㔀㔀⤀ Ⰰഀഀ
	[TZOffset] [smallint] NULL ,਍ऀ嬀唀猀攀爀倀攀爀洀䴀愀猀欀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[RoleFK] [int] NULL ,਍ऀ嬀伀昀昀椀挀攀倀栀漀渀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[HomePhone] [nvarchar] (32) ,਍ऀ嬀䴀漀戀椀氀攀倀栀漀渀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[JobTitle] [nvarchar] (64) ,਍ऀ嬀䨀漀戀䘀甀渀挀琀椀漀渀崀 嬀渀琀攀砀琀崀 Ⰰഀഀ
	[Resume] [ntext] ,਍ऀ嬀䔀䴀愀椀氀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㘀㐀⤀ Ⰰഀഀ
	[PagerEMail] [nvarchar] (64) ,਍ऀ嬀伀昀昀椀挀攀䰀漀挀愀琀椀漀渀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㈀㔀㔀⤀ Ⰰഀഀ
	[PhotoFileFK] [int] NULL ,਍ऀ嬀一漀琀攀猀崀 嬀渀琀攀砀琀崀 Ⰰഀഀ
	[Created] [datetime] NULL ,਍ऀ嬀䰀愀猀琀䄀挀挀攀猀猀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 ഀഀ
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]਍䜀伀ഀഀ
਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䐀攀瀀愀爀琀洀攀渀琀猀崀 ⠀ഀഀ
	[DeptPK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀伀爀最䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[IsActive] [bit] NOT NULL ,਍ऀ嬀䐀攀瀀琀一愀洀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[DeptDesc] [nvarchar] (128) ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblEMailMsgs] (਍ऀ嬀䔀䴀愀椀氀䴀猀最倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[LangFK] [int] NULL ,਍ऀ嬀䔀䴀愀椀氀䴀猀最吀礀瀀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[Subject] [nvarchar] (128) ,਍ऀ嬀䈀漀搀礀崀 嬀渀琀攀砀琀崀 Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[IsActive] [bit] NOT NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀 吀䔀堀吀䤀䴀䄀䜀䔀开伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblFiles] (਍ऀ嬀䘀椀氀攀倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[CaseFK] [nvarchar] (32) ,਍ऀ嬀䘀椀氀攀一愀洀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㈀㔀㔀⤀ Ⰰഀഀ
	[FileLocation] [nvarchar] (128) ,਍ऀ嬀䘀椀氀攀匀椀稀攀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[FileData] [image] NULL ,਍ऀ嬀䌀漀渀琀攀渀琀吀礀瀀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㘀㐀⤀ Ⰰഀഀ
	[UploadDate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀 吀䔀堀吀䤀䴀䄀䜀䔀开伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblGroupMembers] (਍ऀ嬀䜀爀漀甀瀀䴀攀洀倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[GroupFK] [int] NULL ,਍ऀ嬀䌀漀渀琀愀挀琀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[AssignOrder] [int] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblGroups] (਍ऀ嬀䜀爀漀甀瀀倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[GroupName] [nvarchar] (32) ,਍ऀ嬀䜀爀漀甀瀀䐀攀猀挀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㄀㈀㠀⤀ Ⰰഀഀ
	[IsActive] [bit] NOT NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblKnowledgebase] (਍ऀ嬀䬀渀漀眀氀攀搀最攀戀愀猀攀倀䬀崀 嬀戀椀最椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[Issue] [text] ,਍ऀ嬀䌀愀甀猀攀崀 嬀琀攀砀琀崀 Ⰰഀഀ
	[Resolution] [text] ,਍ऀ嬀䔀渀琀攀爀攀搀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[EnteredDate] [datetime] NULL ,਍ऀ嬀䤀猀䄀挀琀椀瘀攀崀 嬀戀椀琀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 ഀഀ
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]਍䜀伀ഀഀ
਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䰀愀渀最甀愀最攀䰀愀戀攀氀猀崀 ⠀ഀഀ
	[LangLabelPK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀䰀愀渀最䰀愀戀攀氀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㘀㐀⤀ Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 ഀഀ
) ON [PRIMARY]਍䜀伀ഀഀ
਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䰀愀渀最甀愀最攀吀攀砀琀猀崀 ⠀ഀഀ
	[LangTextPK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀䰀愀渀最䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[LangLabelFK] [int] NULL ,਍ऀ嬀䰀愀渀最吀攀砀琀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㘀㐀⤀ Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 ഀഀ
) ON [PRIMARY]਍䜀伀ഀഀ
਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䰀愀渀最甀愀最攀猀崀 ⠀ഀഀ
	[LangPK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀䰀愀渀最一愀洀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[Localized] [nvarchar] (32) ,਍ऀ嬀䤀猀刀吀䰀崀 嬀戀椀琀崀 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[Encoding] [nvarchar] (16) ,਍ऀ嬀䤀匀伀㘀㌀㤀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㈀⤀ Ⰰഀഀ
	[IsActive] [bit] NOT NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblLists] (਍ऀ嬀䰀椀猀琀䤀琀攀洀倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[ParentListItemFK] [int] NULL ,਍ऀ嬀䤀琀攀洀伀爀搀攀爀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[IsActive] [bit] NOT NULL ,਍ऀ嬀䤀琀攀洀一愀洀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 ഀഀ
) ON [PRIMARY]਍䜀伀ഀഀ
਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀一漀琀攀猀崀 ⠀ഀഀ
	[NotePK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀䌀愀猀攀䘀䬀崀 嬀戀椀最椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[Note] [ntext] ,਍ऀ嬀䄀搀搀䐀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[IsPrivate] [bit] NOT NULL ,਍ऀ嬀䴀椀渀甀琀攀猀匀瀀攀渀琀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[OwnerFK] [int] NULL ,਍ऀ嬀䈀椀氀氀吀礀瀀攀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 ഀഀ
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]਍䜀伀ഀഀ
਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀伀爀最愀渀椀猀愀琀椀漀渀猀崀 ⠀ഀഀ
	[OrgPK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀伀爀最吀礀瀀攀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[OrgShortName] [nvarchar] (32) ,਍ऀ嬀伀爀最一愀洀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㘀㐀⤀ Ⰰഀഀ
	[IsActive] [bit] NOT NULL ,਍ऀ嬀倀圀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[PrimaryContactFK] [int] NULL ,਍ऀ嬀伀昀昀椀挀攀䰀漀挀愀琀椀漀渀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㈀㔀㔀⤀ Ⰰഀഀ
	[Phone] [nvarchar] (32) ,਍ऀ嬀䘀愀砀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[Email] [nvarchar] (32) ,਍ऀ嬀䴀愀椀氀䄀搀搀爀攀猀猀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㈀㔀㔀⤀ Ⰰഀഀ
	[CourierAddress] [nvarchar] (255) ,਍ऀ嬀䌀椀琀礀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[State] [nvarchar] (32) ,਍ऀ嬀䌀漀甀渀琀爀礀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[Notes] [ntext] ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀 吀䔀堀吀䤀䴀䄀䜀䔀开伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
CREATE TABLE [dbo].[tblParameters] (਍ऀ嬀倀愀爀愀洀倀䬀崀 嬀椀渀琀崀 䤀䐀䔀一吀䤀吀夀 ⠀㄀Ⰰ ㄀⤀ 一伀吀 一唀䰀䰀 Ⰰഀഀ
	[ParamName] [nvarchar] (32) ,਍ऀ嬀倀愀爀愀洀嘀愀氀甀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㄀㈀㠀⤀ Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 ഀഀ
) ON [PRIMARY]਍䜀伀ഀഀ
਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀倀攀爀洀椀猀猀椀漀渀猀崀 ⠀ഀഀ
	[PermPK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀倀攀爀洀䈀礀琀攀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[PermLabel] [nvarchar] (32) ,਍ऀ嬀倀攀爀洀䐀攀猀挀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㄀㈀㠀⤀ Ⰰഀഀ
	[LastUpdate] [datetime] NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀 嬀椀渀琀崀 一唀䰀䰀 ഀഀ
) ON [PRIMARY]਍䜀伀ഀഀ
਍䌀刀䔀䄀吀䔀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀刀漀氀攀猀崀 ⠀ഀഀ
	[RolePK] [int] IDENTITY (1, 1) NOT NULL ,਍ऀ嬀刀漀氀攀一愀洀攀崀 嬀渀瘀愀爀挀栀愀爀崀 ⠀㌀㈀⤀ Ⰰഀഀ
	[RoleDesc] [nvarchar] (128) ,਍ऀ嬀刀漀氀攀䴀愀猀欀崀 嬀椀渀琀崀 一唀䰀䰀 Ⰰഀഀ
	[IsActive] [bit] NOT NULL ,਍ऀ嬀䰀愀猀琀唀瀀搀愀琀攀崀 嬀搀愀琀攀琀椀洀攀崀 一唀䰀䰀 Ⰰഀഀ
	[LastUpdateByFK] [int] NULL ਍⤀ 伀一 嬀倀刀䤀䴀䄀刀夀崀ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblAssignments] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀䄀猀猀椀最渀洀攀渀琀猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  一伀一䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀䄀猀猀椀最渀洀攀渀琀倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䌀愀猀攀吀礀瀀攀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblCaseTypes] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[CaseTypePK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblCases] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀䌀愀猀攀猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  一伀一䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀䌀愀猀攀倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䌀愀琀攀最漀爀椀攀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblCategories] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[CatPK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblContacts] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀䌀漀渀琀愀挀琀猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  一伀一䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䐀攀瀀愀爀琀洀攀渀琀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblDepts] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[DeptPK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblEMailMsgs] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀䔀䴀愀椀氀䴀猀最猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  一伀一䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀䔀䴀愀椀氀䴀猀最倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䘀椀氀攀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblFiles] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[FilePK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblGroupMembers] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀䜀爀漀甀瀀䴀攀洀戀攀爀猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  一伀一䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀䜀爀漀甀瀀䴀攀洀倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䜀爀漀甀瀀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblGroups] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[GroupPK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblKnowledgebase] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀䬀渀漀眀氀攀搀最攀䈀愀猀攀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀䬀渀漀眀氀攀搀最攀戀愀猀攀倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䰀愀渀最甀愀最攀䰀愀戀攀氀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblLangLabels] PRIMARY KEY  CLUSTERED ਍ऀ⠀ഀഀ
		[LangLabelPK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblLanguageTexts] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀䰀愀渀最吀攀砀琀猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀䰀愀渀最吀攀砀琀倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䰀愀渀最甀愀最攀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblLanguages] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[LangPK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblLists] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀䰀椀猀琀猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  一伀一䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀䰀椀猀琀䤀琀攀洀倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀一漀琀攀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblNotes] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[NotePK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblOrganisations] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀伀爀最猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  一伀一䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀伀爀最倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀倀愀爀愀洀攀琀攀爀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblParams] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[ParamPK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblPermissions] WITH NOCHECK ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀倀䬀开琀戀氀倀攀爀洀猀崀 倀刀䤀䴀䄀刀夀 䬀䔀夀  一伀一䌀䰀唀匀吀䔀刀䔀䐀 ഀഀ
	(਍ऀऀ嬀倀攀爀洀倀䬀崀ഀഀ
	)  ON [PRIMARY] ਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀刀漀氀攀猀崀 圀䤀吀䠀 一伀䌀䠀䔀䌀䬀 䄀䐀䐀 ഀഀ
	CONSTRAINT [PK_tblRoles] PRIMARY KEY  NONCLUSTERED ਍ऀ⠀ഀഀ
		[RolePK]਍ऀ⤀  伀一 嬀倀刀䤀䴀䄀刀夀崀 ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblAssignments] ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䄀猀猀椀最渀洀攀渀琀猀开琀戀氀䌀愀猀攀吀礀瀀攀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䌀愀猀攀吀礀瀀攀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblCaseTypes] (਍ऀऀ嬀䌀愀猀攀吀礀瀀攀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䄀猀猀椀最渀洀攀渀琀猀开琀戀氀䌀愀琀攀最漀爀椀攀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䌀愀琀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblCategories] (਍ऀऀ嬀䌀愀琀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䄀猀猀椀最渀洀攀渀琀猀开琀戀氀䌀漀渀琀愀挀琀猀开䰀愀猀琀唀瀀搀愀琀攀䈀礀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䄀猀猀椀最渀洀攀渀琀猀开琀戀氀䌀漀渀琀愀挀琀猀开刀攀瀀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀刀攀瀀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	)਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䌀愀猀攀猀崀 䄀䐀䐀 ഀഀ
	CONSTRAINT [FK_tblCases_tblCaseTypes] FOREIGN KEY ਍ऀ⠀ഀഀ
		[CaseTypeFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀愀猀攀吀礀瀀攀猀崀 ⠀ഀഀ
		[CaseTypePK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblCategories] FOREIGN KEY ਍ऀ⠀ഀഀ
		[CatFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀愀琀攀最漀爀椀攀猀崀 ⠀ഀഀ
		[CatPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblContacts_EnteredBy] FOREIGN KEY ਍ऀ⠀ഀഀ
		[EnteredByFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblContacts_LastUpdateBy] FOREIGN KEY ਍ऀ⠀ഀഀ
		[LastUpdateByFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblContacts_Rep] FOREIGN KEY ਍ऀ⠀ഀഀ
		[RepFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblContacts_User] FOREIGN KEY ਍ऀ⠀ഀഀ
		[ContactFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblDepts] FOREIGN KEY ਍ऀ⠀ഀഀ
		[DeptFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䐀攀瀀愀爀琀洀攀渀琀猀崀 ⠀ഀഀ
		[DeptPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblGroups] FOREIGN KEY ਍ऀ⠀ഀഀ
		[GroupFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䜀爀漀甀瀀猀崀 ⠀ഀഀ
		[GroupPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblLists_Priority] FOREIGN KEY ਍ऀ⠀ഀഀ
		[PriorityFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䰀椀猀琀猀崀 ⠀ഀഀ
		[ListItemPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblCases_tblLists_Status] FOREIGN KEY ਍ऀ⠀ഀഀ
		[StatusFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䰀椀猀琀猀崀 ⠀ഀഀ
		[ListItemPK]਍ऀ⤀ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblCategories] ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䌀愀琀攀最漀爀椀攀猀开琀戀氀䌀愀猀攀吀礀瀀攀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䌀愀猀攀吀礀瀀攀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblCaseTypes] (਍ऀऀ嬀䌀愀猀攀吀礀瀀攀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䌀愀琀攀最漀爀椀攀猀开琀戀氀䌀漀渀琀愀挀琀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	)਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 䄀䐀䐀 ഀഀ
	CONSTRAINT [FK_tblContacts_tblContacts] FOREIGN KEY ਍ऀ⠀ഀഀ
		[LastUpdateByFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblContacts_tblDepts] FOREIGN KEY ਍ऀ⠀ഀഀ
		[DeptFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䐀攀瀀愀爀琀洀攀渀琀猀崀 ⠀ഀഀ
		[DeptPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblContacts_tblFiles] FOREIGN KEY ਍ऀ⠀ഀഀ
		[PhotoFileFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䘀椀氀攀猀崀 ⠀ഀഀ
		[FilePK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblContacts_tblLanguages] FOREIGN KEY ਍ऀ⠀ഀഀ
		[LangFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䰀愀渀最甀愀最攀猀崀 ⠀ഀഀ
		[LangPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblContacts_tblLists_ContactType] FOREIGN KEY ਍ऀ⠀ഀഀ
		[ContactTypeFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䰀椀猀琀猀崀 ⠀ഀഀ
		[ListItemPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblContacts_tblLists_IOStatus] FOREIGN KEY ਍ऀ⠀ഀഀ
		[IOStatusFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䰀椀猀琀猀崀 ⠀ഀഀ
		[ListItemPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblContacts_tblOrgs] FOREIGN KEY ਍ऀ⠀ഀഀ
		[OrgFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀伀爀最愀渀椀猀愀琀椀漀渀猀崀 ⠀ഀഀ
		[OrgPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblContacts_tblRoles] FOREIGN KEY ਍ऀ⠀ഀഀ
		[RoleFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀刀漀氀攀猀崀 ⠀ഀഀ
		[RolePK]਍ऀ⤀ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblDepartments] ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䐀攀瀀琀猀开琀戀氀䌀漀渀琀愀挀琀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䐀攀瀀琀猀开琀戀氀伀爀最猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀伀爀最䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblOrganisations] (਍ऀऀ嬀伀爀最倀䬀崀ഀഀ
	)਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䘀椀氀攀猀崀 䄀䐀䐀 ഀഀ
	CONSTRAINT [FK_tblFiles_tblContacts] FOREIGN KEY ਍ऀ⠀ഀഀ
		[LastUpdateByFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblGroupMembers] ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䜀爀漀甀瀀䴀攀洀戀攀爀猀开琀戀氀䌀漀渀琀愀挀琀猀开䌀漀渀琀愀挀琀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䌀漀渀琀愀挀琀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䜀爀漀甀瀀䴀攀洀戀攀爀猀开琀戀氀䌀漀渀琀愀挀琀猀开䰀愀猀琀唀瀀搀愀琀攀䈀礀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䜀爀漀甀瀀䴀攀洀戀攀爀猀开琀戀氀䜀爀漀甀瀀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䜀爀漀甀瀀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblGroups] (਍ऀऀ嬀䜀爀漀甀瀀倀䬀崀ഀഀ
	)਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䜀爀漀甀瀀猀崀 䄀䐀䐀 ഀഀ
	CONSTRAINT [FK_tblGroups_tblContacts] FOREIGN KEY ਍ऀ⠀ഀഀ
		[LastUpdateByFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblLanguageTexts] ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䰀愀渀最吀攀砀琀猀开琀戀氀䰀愀渀最䰀愀戀攀氀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䰀愀渀最䰀愀戀攀氀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblLanguageLabels] (਍ऀऀ嬀䰀愀渀最䰀愀戀攀氀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀䰀愀渀最吀攀砀琀猀开琀戀氀䰀愀渀最甀愀最攀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䰀愀渀最䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblLanguages] (਍ऀऀ嬀䰀愀渀最倀䬀崀ഀഀ
	)਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀䰀椀猀琀猀崀 䄀䐀䐀 ഀഀ
	CONSTRAINT [FK_tblLists_tblContacts] FOREIGN KEY ਍ऀ⠀ഀഀ
		[LastUpdateByFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblLists_tblLists] FOREIGN KEY ਍ऀ⠀ഀഀ
		[ParentListItemFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䰀椀猀琀猀崀 ⠀ഀഀ
		[ListItemPK]਍ऀ⤀ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblNotes] ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀一漀琀攀猀开琀戀氀䌀愀猀攀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䌀愀猀攀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblCases] (਍ऀऀ嬀䌀愀猀攀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀一漀琀攀猀开琀戀氀䌀漀渀琀愀挀琀猀开䰀愀猀琀唀瀀搀愀琀攀䈀礀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀一漀琀攀猀开琀戀氀䌀漀渀琀愀挀琀猀开伀眀渀攀爀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀伀眀渀攀爀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	),਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀一漀琀攀猀开琀戀氀䰀椀猀琀猀开䈀椀氀氀吀礀瀀攀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䈀椀氀氀吀礀瀀攀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblLists] (਍ऀऀ嬀䰀椀猀琀䤀琀攀洀倀䬀崀ഀഀ
	)਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀伀爀最愀渀椀猀愀琀椀漀渀猀崀 䄀䐀䐀 ഀഀ
	CONSTRAINT [FK_tblOrgs_tblContacts] FOREIGN KEY ਍ऀ⠀ഀഀ
		[LastUpdateByFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblOrgs_tblContacts_PrimaryContact] FOREIGN KEY ਍ऀ⠀ഀഀ
		[PrimaryContactFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀Ⰰഀഀ
	CONSTRAINT [FK_tblOrgs_tblLists] FOREIGN KEY ਍ऀ⠀ഀഀ
		[OrgTypeFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䰀椀猀琀猀崀 ⠀ഀഀ
		[ListItemPK]਍ऀ⤀ഀഀ
GO਍ഀഀ
ALTER TABLE [dbo].[tblPermissions] ADD ਍ऀ䌀伀一匀吀刀䄀䤀一吀 嬀䘀䬀开琀戀氀倀攀爀洀猀开琀戀氀䌀漀渀琀愀挀琀猀崀 䘀伀刀䔀䤀䜀一 䬀䔀夀 ഀഀ
	(਍ऀऀ嬀䰀愀猀琀唀瀀搀愀琀攀䈀礀䘀䬀崀ഀഀ
	) REFERENCES [dbo].[tblContacts] (਍ऀऀ嬀䌀漀渀琀愀挀琀倀䬀崀ഀഀ
	)਍䜀伀ഀഀ
਍䄀䰀吀䔀刀 吀䄀䈀䰀䔀 嬀搀戀漀崀⸀嬀琀戀氀刀漀氀攀猀崀 䄀䐀䐀 ഀഀ
	CONSTRAINT [FK_tblRoles_tblContacts] FOREIGN KEY ਍ऀ⠀ഀഀ
		[LastUpdateByFK]਍ऀ⤀ 刀䔀䘀䔀刀䔀一䌀䔀匀 嬀搀戀漀崀⸀嬀琀戀氀䌀漀渀琀愀挀琀猀崀 ⠀ഀഀ
		[ContactPK]਍ऀ⤀ഀഀ
GO਍ഀഀ
