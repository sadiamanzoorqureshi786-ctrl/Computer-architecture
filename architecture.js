const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, AlignmentType, PageBreak } = require("docx");
const fs = require("fs");

// Dark Green Color Palette Constants
const COLORS = {
  PRIMARY_DARK: "004D40",    // Deep Forest Green for H1 & Main Title
  PRIMARY_MEDIUM: "0E6251",  // Medium Dark Green for H2
  ACCENT_MUTED: "117A65",    // Dark Teal/Green for Notes & Accents
  TBL_HEADER_BG: "E8F8F5",   // Light Mint Green for Table Headers
  CODE_BG: "F2F4F4",         // Light Neutral Gray for Code Blocks
  BORDER_GRAY: "BDC3C7",     // Light Gray for Table/HR Borders
  TEXT_MAIN: "2C3E50",       // Dark Slate Text
  TEXT_MUTED: "566573"      // Subtitle Gray
};

function H1(text) {
  return new Paragraph({ 
    children: [new TextRun({ text, bold: true, color: COLORS.PRIMARY_DARK, size: 30 })], 
    heading: HeadingLevel.HEADING_1, 
    spacing: { before: 360, after: 140 } 
  });
}

function H2(text) {
  return new Paragraph({ 
    children: [new TextRun({ text, bold: true, color: COLORS.PRIMARY_MEDIUM, size: 24 })], 
    heading: HeadingLevel.HEADING_2, 
    spacing: { before: 240, after: 100 } 
  });
}

function P(text, opts = {}) {
  return new Paragraph({ 
    children: [new TextRun({ text, color: COLORS.TEXT_MAIN, ...opts })], 
    spacing: { after: 120, line: 276 } 
  });
}

function Bullet(text, opts = {}) {
  return new Paragraph({ 
    children: [new TextRun({ text, color: COLORS.TEXT_MAIN, ...opts })], 
    bullet: { level: 0 }, 
    spacing: { after: 80 } 
  });
}

function Mono(text) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Consolas", size: 20, color: "1C2833" })],
    spacing: { before: 100, after: 100 },
    shading: { type: ShadingType.CLEAR, fill: COLORS.CODE_BG },
    indent: { left: 240 },
    border: { left: { color: COLORS.PRIMARY_MEDIUM, space: 4, style: BorderStyle.SINGLE, size: 12 } }
  });
}

function Note(text) {
  return new Paragraph({
    children: [
      new TextRun({ text: "Correction / Note: ", bold: true, color: COLORS.ACCENT_MUTED }), 
      new TextRun({ text, italics: true, color: COLORS.ACCENT_MUTED })
    ],
    spacing: { before: 120, after: 150 },
    border: { left: { color: COLORS.ACCENT_MUTED, space: 6, style: BorderStyle.SINGLE, size: 18 } },
    indent: { left: 240 },
  });
}

function HR() {
  return new Paragraph({ 
    text: "", 
    border: { bottom: { color: COLORS.BORDER_GRAY, space: 1, style: BorderStyle.SINGLE, size: 6 } }, 
    spacing: { before: 150, after: 200 } 
  });
}

function cell(text, opts = {}) {
  return new TableCell({
    width: { size: opts.width || 2500, type: WidthType.DXA },
    shading: opts.header ? { type: ShadingType.CLEAR, fill: COLORS.TBL_HEADER_BG } : undefined,
    children: [
      new Paragraph({ 
        children: [
          new TextRun({ 
            text, 
            bold: !!opts.header, 
            color: opts.header ? COLORS.PRIMARY_DARK : COLORS.TEXT_MAIN 
          })
        ] 
      })
    ],
  });
}

function table(colWidths, rows) {
  return new Table({
    columnWidths: colWidths,
    rows: rows.map((r) => new TableRow({ children: r.map((c, i) => cell(c.text, { header: c.header, width: colWidths[i] })) })),
  });
}

const dFlipTable = table(
  [2000, 2500, 2500],
  [
    [{ text: "CLK", header: true }, { text: "D", header: true }, { text: "Q (next state)", header: true }],
    [{ text: "0 (no clock edge)" }, { text: "X (don't care)" }, { text: "Q (holds previous / memory)" }],
    [{ text: "1 (rising edge)" }, { text: "0" }, { text: "0" }],
    [{ text: "1 (rising edge)" }, { text: "1" }, { text: "1" }],
  ]
);

const decoderSizeTable = table(
  [2500, 2500, 2500],
  [
    [{ text: "Select/Input lines (n)", header: true }, { text: "Decoder type", header: true }, { text: "Outputs (2^n)", header: true }],
    [{ text: "1" }, { text: "1-to-2 decoder" }, { text: "2" }],
    [{ text: "2" }, { text: "2-to-4 decoder" }, { text: "4" }],
    [{ text: "3" }, { text: "3-to-8 decoder" }, { text: "8" }],
    [{ text: "4" }, { text: "4-to-16 decoder" }, { text: "16" }],
  ]
);

const decEncTable = table(
  [2200, 3400, 3400],
  [
    [{ text: "Feature", header: true }, { text: "Decoder", header: true }, { text: "Encoder", header: true }],
    [{ text: "Function" }, { text: "Binary code IN → one active output line" }, { text: "One active input line IN → binary code OUT" }],
    [{ text: "Inputs vs Outputs" }, { text: "Few inputs, many outputs (n → 2^n)" }, { text: "Many inputs, few outputs (2^n → n)" }],
    [{ text: "Relation" }, { text: "Opposite of encoder" }, { text: "Opposite of decoder" }],
    [{ text: "Example use" }, { text: "Memory address decoding" }, { text: "Priority encoding, keyboard input" }],
  ]
);

const shiftRegTable = table(
  [2200, 2800, 2800],
  [
    [{ text: "Type", header: true }, { text: "Input (IP)", header: true }, { text: "Output (Q/P)", header: true }],
    [{ text: "SISO" }, { text: "Serial (1 bit at a time)" }, { text: "Serial (1 bit at a time)" }],
    [{ text: "SIPO" }, { text: "Serial (1 bit at a time)" }, { text: "Parallel (all bits at once)" }],
    [{ text: "PISO" }, { text: "Parallel (all bits at once)" }, { text: "Serial (1 bit at a time)" }],
    [{ text: "PIPO" }, { text: "Parallel (all bits at once)" }, { text: "Parallel (all bits at once)" }],
  ]
);

const muxSelectTable = table(
  [1800, 1800, 1800, 2200],
  [
    [{ text: "S2", header: true }, { text: "S1", header: true }, { text: "S0", header: true }, { text: "Selected Input", header: true }],
    [{ text: "0" }, { text: "0" }, { text: "0" }, { text: "I0" }],
    [{ text: "0" }, { text: "0" }, { text: "1" }, { text: "I1" }],
    [{ text: "0" }, { text: "1" }, { text: "0" }, { text: "I2" }],
    [{ text: "0" }, { text: "1" }, { text: "1" }, { text: "I3" }],
    [{ text: "1" }, { text: "0" }, { text: "0" }, { text: "I4" }],
    [{ text: "1" }, { text: "0" }, { text: "1" }, { text: "I5" }],
    [{ text: "1" }, { text: "1" }, { text: "0" }, { text: "I6" }],
    [{ text: "1" }, { text: "1" }, { text: "1" }, { text: "I7" }],
  ]
);

const doc = new Document({
  sections: [
    {
      properties: { page: { size: { width: 12240, height: 15840 } } },
      children: [
        new Paragraph({ children: [new TextRun({ text: "Computer Architecture", bold: true, size: 44, color: COLORS.PRIMARY_DARK })], alignment: AlignmentType.CENTER, spacing: { after: 60 } }),
        new Paragraph({ children: [new TextRun({ text: "Complete Notes: Digital Logic Components, Registers, Shift Registers & Register Transfer", italics: true, size: 24, color: COLORS.TEXT_MUTED })], alignment: AlignmentType.CENTER, spacing: { after: 400 } }),

        H1("1. Integrated Circuit (IC)"),
        P("An Integrated Circuit (IC) is a miniature electronic circuit fabricated on a single chip of semiconductor material (usually silicon). It combines transistors, resistors, gates, and other components into one small package."),
        P("Every IC needs two basic connections:"),
        Bullet("Power (Vcc) — supplies voltage to run the internal circuit"),
        Bullet("Ground (GND) — completes the circuit / reference (0V) line"),
        P("ICs are classified by the number of components they contain: SSI (Small), MSI (Medium), LSI (Large), VLSI (Very Large), and ULSI (Ultra Large Scale Integration). Modern processors are ULSI chips with billions of transistors."),
        HR(),

        H1("2. Decoder and its Types"),
        P("A decoder is a combinational circuit that takes n binary inputs and activates exactly one of 2^n possible outputs — few inputs, more outputs."),
        Mono("Rule:  if number of inputs = n,  then number of outputs = 2^n"),
        H2("2.1 Types of Decoders"),
        decoderSizeTable,
        new Paragraph({ text: "", spacing: { after: 150 } }),
        H2("2.2 Example"),
        P("A 2-to-4 decoder has 2 input lines and 4 output lines. If the input is 10, only output line 2 (counting from 0) becomes HIGH (1); all others stay LOW (0)."),
        H2("2.3 Applications"),
        Bullet("Memory address decoding"),
        Bullet("Instruction decoding in the CPU"),
        Bullet("Device/peripheral selection"),
        HR(),

        H1("3. Encoder"),
        P("An encoder performs the opposite operation of a decoder: it takes many input lines (only one active at a time) and produces a smaller binary code representing which input was active — many inputs, few outputs."),
        Mono("Rule:  2^n inputs  →  n-bit binary output"),
        P("If two or more inputs become active simultaneously, a simple encoder gives wrong/ambiguous output. This is solved by a Priority Encoder, which always encodes the highest-priority active input and ignores the rest."),
        decEncTable,
        new Paragraph({ text: "", spacing: { after: 150 } }),
        HR(),

        H1("4. Multiplexer (MUX)"),
        P("A multiplexer is a combinational circuit with n select lines, 2^n input lines, and only ONE output line. Based on the binary value placed on the select lines, the multiplexer decides which single input line gets connected to the output."),
        Mono("Rule:  n select lines  →  2^n input lines  →  1 output line"),
        H2("4.1 Example: 8-to-1 Multiplexer (8×1 MUX)"),
        P("An 8×1 MUX has 8 data inputs (I0–I7), 3 select lines (S2, S1, S0 — since 2^3 = 8), and 1 output. The select-line combination below decides which input is passed to the output:"),
        muxSelectTable,
        new Paragraph({ text: "", spacing: { after: 150 } }),
        Note("In your handwritten notes the select-line weights were written as 1, 2, 4, 8 read right-to-left (i.e., S0=1, S1=2, S2=4/8-weighting). The standard convention is binary weighting S2S1S0 = 4-2-1, matching the table above; I've corrected the ordering here for clarity."),
        P("Multiplexers are built internally using AND, OR, and NOT gates and are widely used in CPUs, communication systems, and ALUs to reduce wiring by sharing one output line among many possible inputs."),
        HR(),

        H1("5. D Flip-Flop"),
        P("A D (Data/Delay) Flip-Flop is the basic memory element in digital circuits. It stores 1 bit of data and is controlled by a Clock (CLK) signal."),
        Bullet("D = data input (the bit to be stored)"),
        Bullet("CLK = clock input (controls when the stored value updates)"),
        Bullet("Q = output (the currently stored/memorized bit)"),
        P("Working: when the clock is inactive (0), the flip-flop ignores the D input and simply holds/remembers its last stored value (memory). When an active clock edge (1 / rising edge) occurs, whatever value is on D at that moment is captured and appears at Q."),
        dFlipTable,
        new Paragraph({ text: "", spacing: { after: 150 } }),
        Note("Your notes listed the first row as CLK=0, D=X, Q=Hold(memory) — this is correct and has been kept as-is; X means 'don't care', since D is ignored while the clock is not active."),
        HR(),

        H1("6. Registers"),
        P("A register is a group of flip-flops connected together and used to store binary information (multiple bits) temporarily inside a digital system. Since 1 flip-flop stores 1 bit, an n-bit register requires n flip-flops, all sharing a common clock so they update together."),
        P("Registers are among the fastest storage elements in a computer, much faster than main memory, and are used extensively inside the CPU (e.g., holding data, addresses, instructions, and intermediate results during program execution)."),
        H2("6.1 Why Registers Are Needed"),
        P("Accessing main memory for every single operation would be very slow. Registers sit right next to the processing circuitry and provide very fast temporary storage, which is essential for high processor performance."),
        H2("6.2 Registers with Parallel Load"),
        P("A parallel-load register loads all its bits at the same time, in a single clock pulse, using an extra control signal called Load:"),
        Bullet("Load = 1 (active): new input data is stored into the register on the clock edge"),
        Bullet("Load = 0 (inactive): the register keeps/retains its previous value"),
        P("Example applications of registers: accumulator register, instruction register (IR), memory address register (AR/AR), data register (DR), program counter (PC), and general temporary storage."),
        HR(),

        H1("7. Shift Registers (Left & Right Shift)"),
        P("A shift register is a special register built from flip-flops connected in a chain (cascade), where data moves ('shifts') from one flip-flop to the next on every clock pulse, either to the left or to the right."),
        H2("7.1 Right Shift — Example"),
        P("Starting value in a 4-bit shift register: 0 1 1 0 (decimal 6). Shifting right by one position moves every bit one place to the right, and the rightmost bit is shifted out:"),
        Mono("0 1 1 0  (=6)  --right shift-->  0 0 1 1  (=3)"),
        P("Notice that shifting right by 1 position is the same as dividing the value by 2: 6 ÷ 2 = 3."),
        H2("7.2 Left Shift — Example"),
        P("Taking 0 0 1 1 (=3) and shifting left by one position moves every bit one place to the left, and a 0 fills in on the right:"),
        Mono("0 0 1 1  (=3)  --left shift-->  0 1 1 0  (=6)"),
        P("Shifting left by 1 position is the same as multiplying the value by 2: 3 × 2 = 6."),
        Note("Your notes wrote '6 ÷ 3 = 2' and '3 × 2 = 6' next to the shift diagrams — the correct arithmetic relationship for a 1-bit shift is division/multiplication by 2 (not 3), i.e. 6 ÷ 2 = 3 for the right shift, and 3 × 2 = 6 for the left shift. I've corrected this above."),
        HR(),

        H1("8. Types of Shift Registers (SISO, SIPO, PISO, PIPO)"),
        P("Shift registers are classified by how data enters and leaves them — serially (one bit at a time) or in parallel (all bits at once):"),
        Bullet("SISO (Serial-In, Serial-Out): data enters one bit at a time and leaves one bit at a time."),
        Bullet("SIPO (Serial-In, Parallel-Out): data enters one bit at a time but all stored bits are read out together at once."),
        Bullet("PISO (Parallel-In, Serial-Out): all bits are loaded into the register at once, then read out one bit at a time."),
        Bullet("PIPO (Parallel-In, Parallel-Out): all bits are loaded in at once and read out all at once."),
        shiftRegTable,
        new Paragraph({ text: "", spacing: { after: 150 } }),
        Note("Your notes' input/output table used 'n' and 'n-1' to describe timing (an n-bit register takes n clock pulses to shift fully in serial mode, and the output appears after the n-1'th shift). The table above restates the same idea in plain serial-vs-parallel terms so it's easier to remember for exams."),
        HR(),

        H1("9. Bidirectional Shift Registers"),
        P("A bidirectional shift register can shift stored data in either direction — left or right — under the control of a mode-select input."),
        Bullet("Mode = 0 → Left shift"),
        Bullet("Mode = 1 → Right shift"),
        P("A bidirectional shift register with parallel load additionally supports: loading all bits at once (parallel load), shifting left, shifting right, or holding its current value — with a small set of select lines choosing which operation happens on the next clock pulse. This combined design is called a Universal Shift Register and is used heavily in arithmetic and data-communication circuits."),
        H2("9.1 Duty Cycle (related clock concept)"),
        P("A duty cycle is the fraction of time, during one full clock cycle, in which the clock signal stays at logic 1 (HIGH) versus logic 0 (LOW). It describes what portion of each cycle the signal is 'on'."),
        HR(),

        H1("10. Register Transfer / Micro-operation"),
        P("A register transfer (or micro-operation) is the movement of binary data from one register to another inside the CPU, during a single clock pulse."),
        P("Main components involved in register transfer:"),
        Bullet("1. ALU (Arithmetic Logic Unit) — performs calculations and logic operations"),
        Bullet("2. Memory — stores data and instructions"),
        Bullet("3. Registers — fast temporary storage locations"),
        P("The Control Unit (CU) is considered the 'brain' of the CPU, and registers are used because they are extremely fast compared to memory."),
        H2("10.1 Register Transfer Language (RTL)"),
        P("Register Transfer Language (RTL) is a symbolic (shorthand) notation used to describe register operations clearly, using an arrow (←) to mean 'is replaced by' / 'gets the value of'."),
        Mono("R2 ← R1     (means: the value of R1 is copied into R2)"),
        H2("10.2 Conditional Register Transfer"),
        P("A conditional register transfer only happens if a stated condition (usually a control signal, P) is true. It is written with the condition before a colon:"),
        Mono("P : R2 ← R1     (means: IF P = 1, THEN copy R1 into R2)"),
        HR(),

        H1("11. Types of Register Transfer"),
        P("There are four main types of register transfer / micro-operations:"),
        H2("11.1 Register Transfer (simple copy)"),
        Mono("R1 ← R2     (copies contents of R2 into R1, R2 is unchanged)"),
        H2("11.2 Arithmetic Micro-operation"),
        Mono("R2 ← R1 + R2     (adds R1 and R2, stores the result in R2)"),
        P("Worked example: If R1 = 18 and R2 = 20, and we perform R2 ← R1 + R2:"),
        Mono("R1 + R2 = 18 + 20 = 38   →   R2 = 38 (after the operation)"),
        H2("11.3 Logic Micro-operation"),
        Mono("R1 ← R1 AND R2     (bitwise AND between R1 and R2, result stored in R1)"),
        H2("11.4 Shift Micro-operation"),
        Mono("R2 ← shift R2     (shifts the bits of R2 left or right by one position)"),
        H2("11.5 Worked Assignment Example"),
        P("Given: R1 = 11110000, R2 = 10101010, R3 = 00110011. Perform in order:"),
        Bullet("R3 ← R1   →  R3 becomes 11110000 (copy of R1)"),
        Bullet("R2 ← R3   →  R2 becomes the (new) value of R3, i.e. 11110000"),
        Bullet("R1 ← R2   →  R1 becomes the (new) value of R2"),
        Bullet("If P = 1, perform R3 ← R1 → only executes when the condition P is true"),
        Note("This is a conditional + sequential register transfer example from your assignment — remember that each transfer uses the value of the source register at that point in time (i.e., after any earlier transfer in the same sequence has already happened), not the original starting value."),
        HR(),

        H1("12. Register Transfer Using Bus"),
        P("Instead of wiring every register directly to every other register (which needs a huge number of wires), all registers can be connected to a single shared communication line called a bus."),
        P("In this technique, if any register wants to transfer data to another register, it can be done through the shared bus: the sending register places its value onto the bus, and the receiving register reads the value off the bus."),
        Mono("R1 ← Bus        Bus ← R2        (i.e. R1 gets whatever value R2 placed on the bus)"),
        P("This greatly reduces wiring complexity, since one bus can be time-shared by many registers instead of needing a separate dedicated path between every possible pair of registers."),
        HR(),

        H1("13. Three-State (Tri-State) Bus Buffers"),
        P("A three-state (tri-state) buffer is a special gate placed between a register and the bus. Unlike a normal digital signal (which only has 2 states, 0 and 1), a tri-state buffer supports 3 possible states:"),
        Bullet("0 — logic LOW output"),
        Bullet("1 — logic HIGH output"),
        Bullet("Z (High Impedance) — the output is electrically disconnected from the bus, as if not connected at all"),
        P("The high-impedance (Z) state means the device is disconnected from the bus, so it does not interfere with whatever value other devices are placing on the bus at that time. This allows many registers to share the same bus safely — only the one currently 'enabled' actually drives the bus, and all the others go into their high-impedance (Z) state."),
        HR(),

        H1("14. Bus Contention"),
        P("Bus contention occurs when more than one device tries to drive (place a value on) the bus at the same time, which can cause conflicting/undefined signal levels and possible damage to the circuit."),
        P("To avoid this, only ONE enable signal is allowed to be active (1) at any given time; every other device connected to the bus must have its enable signal at 0 (placing it into the high-impedance / disconnected state)."),
        P("Example sequence using a bus safely (only one source enabled at a time):"),
        Mono("R1 ← R2   is done as:   Bus ← R2   (R2 enabled, placed onto bus)\n                        R1 ← Bus  (R1 reads the value from the bus)"),
        Note("Only R2's enable signal should be active for that step; R1 (and any other register connected to the bus) must be disabled/high-impedance so their signals don't clash with R2's on the shared bus line."),
        HR(),

        H1("15. Memory Transfer"),
        P("Memory transfer describes how data moves between the CPU's registers and main memory, using two special registers:"),
        Bullet("AR (Address Register) — holds the memory address being accessed"),
        Bullet("DR (Data Register) — holds the actual data being read from or written to memory"),
        P("The address bus carries the information about WHICH memory location is being accessed; the data bus carries the actual value being transferred."),
        H2("15.1 Memory Read Example"),
        P("To read the value stored at memory location 100 (suppose Mem[100] = 52):"),
        Mono("AR ← 100                (the address 100 is loaded into the Address Register)\nDR ← Mem[100]           (the data at that address, 52, is loaded into the Data Register)"),
        H2("15.2 Memory Write Example"),
        P("To write the value 50 into memory location 100:"),
        Mono("DR ← 50                 (the value to be written, 50, is placed in the Data Register)\nAR ← 100                (the target address, 100, is placed in the Address Register)\nMem[100] ← DR    i.e.   Mem[100] = 50   (the value is written into that memory location)"),
        P("In short: to READ, the address goes into AR, and the data found there is copied out into DR. To WRITE, the data goes into DR, the address goes into AR, and then that data is copied from DR into the memory location pointed to by AR."),
        HR(),

        H1("16. Short Questions (Registers & Shift Registers)"),
        Bullet("1. Define a register."),
        Bullet("2. Why are registers faster than main memory?"),
        Bullet("3. What is a parallel load register?"),
        Bullet("4. How many flip-flops are required for an 8-bit register?"),
        Bullet("5. Define a shift register."),
        Bullet("6. Differentiate between SISO and SIPO shift registers."),
        Bullet("7. What is a bidirectional shift register?"),
        Bullet("8. What is the purpose of the Load control signal?"),
        Bullet("9. Explain left shift and right shift operations."),
        Bullet("10. State two applications of shift registers."),

        H1("17. MCQs (Registers & Shift Registers, with answers)"),
        Bullet("1. An n-bit register requires: A) n gates B) n flip-flops C) n multiplexers D) n decoders → Answer: B"),
        Bullet("2. Which register loads all bits simultaneously? A) Shift Register B) Serial Register C) Parallel Load Register D) Counter → Answer: C"),
        Bullet("3. Which shift register converts serial data into parallel data? A) PISO B) SIPO C) PIPO D) SISO → Answer: B"),
        Bullet("4. A bidirectional shift register can shift: A) Left only B) Right only C) Both directions D) None → Answer: C"),
        Bullet("5. Which component is commonly used to construct registers? A) Decoder B) Encoder C) Flip-Flop D) Multiplexer → Answer: C"),
        HR(),

        H1("18. Quick Revision Summary"),
        Bullet("IC → chip combining components; needs Power + Ground; classified SSI→ULSI."),
        Bullet("Decoder → n inputs → 2^n outputs (binary code to one active line)."),
        Bullet("Encoder → 2^n inputs → n outputs (one active line to binary code); Priority Encoder handles multiple active inputs."),
        Bullet("Multiplexer → n select lines → 2^n inputs → 1 output (data selector)."),
        Bullet("D Flip-Flop → stores 1 bit; updates on clock edge; holds value (memory) otherwise."),
        Bullet("Register → group of flip-flops storing multiple bits; parallel load via Load signal."),
        Bullet("Shift Register → shifts bits left/right each clock pulse; left shift ≈ ×2, right shift ≈ ÷2 (per 1-bit shift)."),
        Bullet("SISO/SIPO/PISO/PIPO → differ by serial vs parallel input and output."),
        Bullet("Bidirectional shift register → Mode 0 = left shift, Mode 1 = right shift."),
        Bullet("Register Transfer (RTL) → R2 ← R1 notation; conditional form P: R2 ← R1."),
        Bullet("4 types of register transfer → simple transfer, arithmetic, logic, shift micro-operations."),
        Bullet("Bus transfer → many registers share one bus instead of direct wiring between every pair."),
        Bullet("Tri-state buffer → 0, 1, and Z (high impedance/disconnected) states."),
        Bullet("Bus contention → avoided by enabling only ONE driver on the bus at a time."),
        Bullet("Memory transfer → AR holds address, DR holds data; read = AR then DR; write = DR and AR then Mem ← DR."),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync('Computer_Architecture_Full_Notes.docx', buffer);
  console.log("done");
});