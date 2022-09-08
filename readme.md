This shares two case study mockups for simulating manufacturing operations.

#### Introduction
Manufacturing operations can be modeled by organizing the inputs affecting the operation and applying calculations to study endpoints. Cost or capacity modeling are examples. Another example is time-based simulations to look at effects of in-market changes in raw material pricing or qualification of new supply points or manufacturing capacity.

For cost modeling, the goal is to understand how cost components add up to a total such as cost per kg of finished goods. By adding company overhead and profit margin inputs, this allows what-if scenarios for profitability. Capacity models use many of the same inputs as cost modeling but can look at how to de-bottleneck a multi-step operation. The **Cost Model.xlsm** file is a mockup showing how this scenario-based model can be organized in Excel. Since project objectives vary, this is a starting point that can be tailored to answer specific objectives.

A second, common modeling objective is to understand the interplay of inputs that change over time. The **TimeSim.xlsx** mockup is a jumping-off point and mockup of how to do this. The mockup purports to look at how a new, supply point's qualification date affects cumulative spending on material supply.

 These models are in Microsoft Excel (Windows client version). We use our [ExcelSteps Addin](https://github.com/jlandgre/ExcelSteps) to curate and refresh model data. The **Why Excel?** section below gives more background on this approach.

#### Operation "Cost" Model Example
Scenario models are very useful for studying operational endpoints such cost per kg of finished goods or line capacity. Each scenario typically depends on dozens of inputs. The example **Cost Model.xlsm** shows a preferred, hub-spokes architecture we use. It divides inputs into three categories: Mixes, Operation and Financial assumptions. Each row on these rows x columns tables is named and can be selected by name on the Model sheet to feed the right inputs to an overall what-if scenario in Model-sheet column.

The example is a starting point for looking at various operational situations such as a daily versus round-the-clock operation schedule or single versus multi-step production sequences. By varying the selected "Mix", it can simulate different products or recipes --including using one scenario's final cost as the raw material cost for a subsequent step in the production chain. Finally, by varying financial assumptions, it can look at questions such as would it make sense to contract manufacture this product?  The mockup's Model-sheet column scenarios show handling of by-weight, by-area and by-parts scenarios. These can be selected as the basis depending on the physical format of the product being modeled.

<p align="center">
 Operation Model Preferred Architecture</br>
 <img src=images/Cost_Model_Block_Diagram.png "Scenario Architecture" width=600></br>
</p>

#### Time-Based Business Simulation Example
In this mockup, a goal might be to look at how the date of a supply qualification such as the Start_Supply_B input on the Qualn sheet affects cumulative raw material spending as in the example graph. Such visuals have proven very helpful for driving collaborative decision-making.

It's an arbitrary choice to put the "model" on the Business worksheet in columns-by-rows scenario model format. Aesthetically, we find it intuitive to have time-based simulations march from left to right across the sheet. The metadata columns at left place calculation formulas and formatting instructions on the same sheet as model calculations. ExcelSteps' Parse Scenario Model menu option allows this to be instantly flipped into rows by columns for graphing and for concatenating multiple simulations to compare them. Note the **Sim_Name** variable in Business sheet Row 5. It can be changed either manually or programatically for each what-if simulation. It serves as an ID when simulations are combined in a graphics and stats package such as [JMP software](https://www.jmp.com/en_us/home.html) used in this demo for plotting data.

<p align="center">
 Time-Based Business Simulation Work Process</br>
 <img src=images/TimeSim_Overview.png "TimeSim Overview" width=700></br>
</p>

#### Why Excel?
These simulations can be done in a variety of software platforms, but there are a couple of reasons Excel makes either a good starting point or a place to do the whole job in an efficient way. First, the combination of Excel and our ExcelSteps add-in makes it easy to add variables and calculations and to validate/curate the formulas. Excel is also pervasive in most companies, so it makes the effort more collaborative. A key task with these models is assembling inputs and keeping them up to date. By being careful to use a modular architecture, functional experts can look after key inputs in their responsibility areas --without having to learn to code.

The ExcelSteps Addin "curates" variables' formulas so that they are refreshed as text strings from validated sources. This is important for models whose outputs make pivotal predictions about plant operations and overall business profitability. With the Addin, the formula curation sources are either the ExcelSteps recipe sheet for rows by columns tables or the metadata columns at left for the columns by rows scenario models. In addition to the important task of curating formulas, ExcelSteps takes care of formatting such as setting number formats, column widths and expandable column outline sections. That's important for allowing model complexity to grow organically as needed while maintaining an intuitive user interface.

<p align="center">
 ExcelSteps Addin Menu and Refresh Dialog Box for TimeSim.xlsx Mockup</br>
 <img src=images/ExcelSteps.png "ExcelSteps" width=600></br>
</p>

Finally, while not [yet?] open-sourced from our consulting practice, for larger projects, we utilize an extensive, Excel/VBA validation suite inspired by the pytest package in the Python world. It uses assertion-based testing and enables continual reverification of Excel calculations in complex models. This avoids the problem of future changes inadvertently breaking previously-verified calculations.

Both of the example models involve a mix of rows/columns data and Scenario Models. If a model needs to progress in complexity, ExcelSteps organization of variable metadata make it easy to transition to Python class objects or other languages. Examples of these progressions might be progressing to simulate stochastic variability to inputs or doing linear programming simulations to optimize product mix, crewing or other simulation outputs.

 J.D. Landgrebe
 Data Delve LLC
