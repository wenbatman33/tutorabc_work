<?Infragistics.Models format="xaml" version="2"?>
<Flow xmlns="http://prototypes.infragistics.com/flows"
                                                         xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
    <Flow.Items>
        <Step x:Uid="$2" ContentView="/Screen2-2.screen" ContentState="9b419995-31d3-4f5b-a7b0-1e7aa2c16a1d" X="1561" Y="108" Width="430" Height="322" />
        <Step ContentView="/Screen1-1.screen" ContentState="03fceea3-2561-4adf-944e-a72a30f58382" X="64" Y="884" Width="430" Height="322" />
        <Step ContentView="/Screen1-2.screen" ContentState="b069d763-bf9e-43eb-a5c2-11dba479efd8" X="-480" Y="884" Width="430" Height="322" />
        <Step x:Uid="$1" ContentView="/Screen2-1.screen" ContentState="20fc6fff-6b06-4d0e-b5fb-4808e40bc073" X="1068" Y="108" Width="430" Height="322" />
        <Connector Source="{Reference $1}" Target="{Reference $2}" />
        <Step x:Uid="$3" ContentView="/Screen2-3.screen" ContentState="42163fff-fe97-4d65-b916-c042d64143ca" X="2041" Y="108" Width="430" Height="322" />
        <Connector Source="{Reference $2}" Target="{Reference $3}" />
        <Step x:Uid="$4" ContentView="/Screen2-4.screen" ContentState="81169267-31ba-411f-ae8f-1a141d9aeeb0" X="2521" Y="108" Width="430" Height="322" />
        <Connector Source="{Reference $3}" Target="{Reference $4}" />
        <Step x:Uid="$5" ContentView="/Client3-1-1.screen" ContentState="739fbb4c-253b-4510-a1d9-9d8ad21ec5fc" ContentSceneHotSpotWidth="277" ContentSceneHotSpotHeight="159" X="1068" Y="1264" Width="430" Height="322" />
        <Step x:Uid="$6" ContentView="/Client3-1-2.screen" ContentState="c96f5457-827d-4cc4-8d93-8d08cb44edee" X="1068" Y="1637" Width="430" Height="322" />
        <Connector Source="{Reference $5}" Target="{Reference $6}" />
        <Step x:Uid="$21" ContentView="/Client0_DashBoard.screen" ContentState="5ac97174-6e85-46e0-b542-67a63ea246ae" ContentSceneHotSpotWidth="277" ContentSceneHotSpotHeight="159" X="64" Y="103" Width="430" Height="322" />
        <Step x:Uid="$7" ContentView="/Client3-1-3.screen" ContentState="887b9e69-d39d-4d04-b127-72169730e1e5" X="1068" Y="2010" Width="430" Height="322" />
        <Connector Source="{Reference $6}" Target="{Reference $7}" />
        <Step x:Uid="$8" ContentView="/Client3-1-4.screen" ContentState="a24fc0e3-ddc6-42b2-909d-e23d4423a6eb" X="1068" Y="2384" Width="430" Height="322" />
        <Connector Source="{Reference $7}" Target="{Reference $8}" />
        <Step x:Uid="$9" ContentView="/Client3-1-5.screen" ContentState="8343dbfe-0a50-4e62-8aac-4b774eaf82a2" X="1057" Y="2765" Width="430" Height="322" />
        <Connector Source="{Reference $8}" Target="{Reference $9}" />
        <Step x:Uid="$10" ContentView="/Client3-1-6.1.screen" ContentState="0a57235e-67bb-4c78-b23a-d94b28c1db06" X="1057" Y="3140" Width="430" Height="322" />
        <Connector Source="{Reference $9}" Target="{Reference $10}" />
        <Step ContentView="/Client3-1-6.2.screen" ContentState="0a7e3b0a-839a-42a8-832c-1840ea908ab4" ContentSceneHotSpotWidth="277" ContentSceneHotSpotHeight="159" X="1550" Y="3140" Width="430" Height="322" />
        <Step x:Uid="$11" ContentView="/Client3-1-7.screen" ContentState="3ca322b0-6342-4ec4-936a-597bb3e8658f" X="1057" Y="3514" Width="430" Height="322" />
        <Connector Source="{Reference $10}" Target="{Reference $11}" />
        <Step x:Uid="$12" ContentView="/Client3-1-8.screen" ContentState="979e5ca6-90b8-4cda-8aeb-4b90a1af2bd3" X="1057" Y="3911" Width="430" Height="322" />
        <Connector Source="{Reference $11}" Target="{Reference $12}" />
        <Step x:Uid="$13" ContentView="/Client3-1-9.screen" ContentState="726132a1-48aa-4b2e-9108-262999e96118" X="1057" Y="4299" Width="430" Height="322" />
        <Connector Source="{Reference $12}" Target="{Reference $13}" />
        <Step x:Uid="$19" ContentView="/Client3-1.screen" ContentState="c580b6f6-1f51-4120-99ef-b91031ad3c24" ContentSceneHotSpotWidth="277" ContentSceneHotSpotHeight="159" X="1068" Y="884" Width="430" Height="322" />
        <Step x:Uid="$14" ContentView="/Client3-1-10.screen" ContentState="101a77f2-86ea-4fc3-bf68-6b8ea67ed3e8" X="1057" Y="4679" Width="430" Height="322" />
        <Connector Source="{Reference $13}" Target="{Reference $14}" />
        <Step x:Uid="$15" ContentView="/Client3-1-11.screen" ContentState="cabe564e-a4b3-4a74-b9da-92897542867f" X="1057" Y="5059" Width="430" Height="322" />
        <Connector Source="{Reference $14}" Target="{Reference $15}" />
        <Step x:Uid="$16" ContentView="/Client3-1-12.screen" ContentState="2d75a047-a9d0-45ca-af2a-8c7eb4e3cc68" X="1057" Y="5449" Width="430" Height="322" />
        <Connector Source="{Reference $15}" Target="{Reference $16}" />
        <Step x:Uid="$17" ContentView="/Client3-1-13.screen" ContentState="b9d51b8d-8699-4421-85a8-a39b2d9e883c" X="1057" Y="5849" Width="430" Height="322" />
        <Connector Source="{Reference $16}" Target="{Reference $17}" />
        <Step x:Uid="$18" ContentView="/Client3-1-14.screen" ContentState="6913c30f-b58f-4417-9637-70053b50ad8a" X="1057" Y="6235" Width="430" Height="322" />
        <Connector Source="{Reference $17}" Target="{Reference $18}" />
        <Step x:Uid="$20" ContentView="/Client3-1.1.screen" ContentState="29970170-7fac-465b-be79-9df3b8770ca5" X="580" Y="1264" Width="430" Height="322" />
        <Connector Source="{Reference $19}" Target="{Reference $20}" />
        <Step x:Uid="$22" ContentView="/Client0-1.screen" ContentState="4149eb45-9016-483a-8214-137dcdf3c3bb" X="556" Y="107" Width="430" Height="322" />
        <Connector Source="{Reference $21}" Target="{Reference $22}" />
        <Step x:Uid="$23" ContentView="/Client0-1.1.screen" ContentState="c32671a5-d203-4871-9172-74c6a44187f7" X="556" Y="493" Width="430" Height="322" />
        <Connector Source="{Reference $22}" Target="{Reference $23}" />
        <Connector Source="{Reference $19}" Target="{Reference $5}" />
        <Step x:Uid="$24" ContentView="/Client3-2.screen" ContentState="711fcde6-5547-40c1-8fe0-78d78e47f8b2" X="1548" Y="884" Width="430" Height="322" />
        <Connector Source="{Reference $19}" Target="{Reference $24}" />
        <Step x:Uid="$25" ContentView="/Client0_DB1.screen" ContentState="2adac240-3340-4c49-9412-cd24413af516" ContentSceneHotSpotWidth="277" ContentSceneHotSpotHeight="159" X="64" Y="-276" Width="430" Height="322" />
        <Step x:Uid="$26" ContentView="/Client0_DB2.screen" ContentState="3badd0e6-b17f-4776-b919-4f89b5fac2d1" X="544" Y="-276" Width="430" Height="322" />
        <Connector Source="{Reference $25}" Target="{Reference $26}" />
        <Step x:Uid="$27" ContentView="/Client0_DB3.screen" ContentState="4a324b9a-6586-4a78-9c9e-fca1e79a193e" X="1024" Y="-276" Width="430" Height="322" />
        <Connector Source="{Reference $26}" Target="{Reference $27}" />
        <Step x:Uid="$28" ContentView="/Client0_DB4.screen" ContentState="983b3425-fd89-482b-9361-4c53af0829d0" X="1504" Y="-276" Width="430" Height="322" />
        <Connector Source="{Reference $27}" Target="{Reference $28}" />
    </Flow.Items>
</Flow>
