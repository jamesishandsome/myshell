import { expect, test } from "bun:test";
import {run} from "../zip-script";



test("test1", async () => {
    const res = await run('./test/test1')
    expect(res?.firstWordBold).toBe(true);
    expect(res?.secondWordUnderline).toBe("single");
    expect(res?.thirdWordFontSize).toBe("40");
});


test("test2", async () => {
    const res = await run('./test/test2')
    expect(res?.firstWordBold).toBe("mixed");
    expect(res?.secondWordUnderline).toBe("mixed");
    expect(res?.thirdWordFontSize).toBe("mixed");
});

test("test3", async () => {
    const res = await run('./test/test3')
    expect(res?.firstWordBold).toBe(true);
    expect(res?.secondWordUnderline).toBe("double");
    expect(res?.thirdWordFontSize).toBe("100");
})

test("test4", async () => {
    const res = await run('./test/test4')
    expect(res?.firstWordBold).toBe(true);
    expect(res?.secondWordUnderline).toBe("double");
    expect(res?.thirdWordFontSize).toBe("100");
})

test("test5", async () => {
    const test5 =async () => {
        const res = await run('./test/test5')
    }
    expect(test5).toThrowError(
        "Not enough words found"
    );
})

