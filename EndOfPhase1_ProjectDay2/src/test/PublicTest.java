package test;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertAll;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import day2.JackOfAll;

@DisplayName("Day 2 Practice Project ")
public class PublicTest {

	JackOfAll jack;

	@BeforeAll
	public static void init() {
		// System.out.println("Test Has Started");

	}

	@BeforeEach
	public void setUp() {
		jack = new JackOfAll();
	}

	@DisplayName("Testing UserName generator")
	@Test
	public void testUserNameGeneration() {

		assertAll(
				() -> assertEquals("ajackson", jack.generateUserName("Adam", "Jackson"), "Wrong Username Generated"),
				() -> assertEquals("dmiller", jack.generateUserName("Diana", "Miller"), "Wrong Username Generated")
				);
	}

	@DisplayName("Testing Reverse Word")
	@Test
	public void testReverseWord() {
		assertAll(
				() -> assertEquals(new StringBuilder("madA").toString(), jack.reverseWords("Adam").toString(),
						"not reversed correctly"),
				() -> assertEquals(new StringBuilder("BBAA321").toString(), jack.reverseWords("123AABB").toString(),
						"not reversed correctly"));
	}

	@DisplayName("Testing Sum of Number from 1 to Given number")
	@Test
	public void testSumOfNumberFrom1ToN() {

		assertAll(() -> assertEquals(6, jack.sumOfNumberFrom1ToN(3), "Sum of the Numbers Wrong"),
				() -> assertEquals(325, jack.sumOfNumberFrom1ToN(25), "Sum of the Numbers Wrong"),
				() -> assertEquals(55, jack.sumOfNumberFrom1ToN(10), "Sum of the Numbers Wrong"));
	}

	@DisplayName("Calculate Shipping Test")
	@Test
	public void testCalculateShipping() {
		assertAll(() -> assertEquals(0, jack.calculateShipping(true, 100)),
				() -> assertEquals(0, jack.calculateShipping(true, 1)),
				() -> assertEquals(0, jack.calculateShipping(false, 36)),
				() -> assertEquals(5, jack.calculateShipping(false, 35)),
				() -> assertEquals(5, jack.calculateShipping(false, 1)));

	}

	@DisplayName("Make First Letter Capital Test")
	@Test
	public void testMakeFirstLetterCapitalTheRestLowerCase() {

		assertAll(() -> assertEquals("Abcd", jack.makeFirstCharCapital("ABCD")),
				() -> assertEquals("Abcde", jack.makeFirstCharCapital("abcdE")),
				() -> assertEquals("Abcdefg", jack.makeFirstCharCapital("abCdEFG")));

	}

	@DisplayName("Make First Character Of Every Word Capital Test")
	@Test
	public void testMakeFirstLetterEveryWordCapitalCapitalTheRestLowerCase() {
		assertAll(() -> assertEquals("Abc Bs Vvd", jack.makeFirstCharEveyWordCapital("Abc bs vvd")),
				() -> assertEquals("Ab Cd E", jack.makeFirstCharEveyWordCapital("ab cd E")),
				() -> assertEquals("Ab Cd Ef Gd", jack.makeFirstCharEveyWordCapital("ab Cd EF Gd")));
	}

	@DisplayName("Get Number Out Of String Test")
	@Test
	public void testGetNumberOutOfString() {
		assertAll(() -> assertEquals(123, jack.getNumberOutOfString("abc1m23")),
				() -> assertEquals(0, jack.getNumberOutOfString("abc")),
				() -> assertEquals(1234, jack.getNumberOutOfString("1234")));
	}

	@DisplayName("Check int Array First Or Last Item more than 5")
	@Test
	public void testCheckIfArrayFirstOrLastItemMoreThan5() {
		assertAll(() -> assertTrue(jack.checkIfFirstOrLastItemMoreThan5(new int[] { 6, 3, 2, 3, 2, 6 })),
				() -> assertTrue(jack.checkIfFirstOrLastItemMoreThan5(new int[] { 1, 2, 6 })),
				() -> assertTrue(jack.checkIfFirstOrLastItemMoreThan5(new int[] { 6, 3, 1 })),
				() -> assertFalse(jack.checkIfFirstOrLastItemMoreThan5(new int[] { 1, 3, 1 })),
				() -> assertFalse(jack.checkIfFirstOrLastItemMoreThan5(new int[] { 1, 2, 3, 1 })));

	}

	@DisplayName("Test Reverse int Array")
	@Test
	public void testReverseIntArray() {

		assertArrayEquals(new int[] { 1, 2, 3, 4, 5 }, jack.reverseIntArray(new int[] { 5, 4, 3, 2, 1 }));
		assertArrayEquals(new int[] { 11, 2, 32, 44, 5 }, jack.reverseIntArray(new int[] { 5, 44, 32, 2, 11 }));

	}

	@DisplayName("Check All Integers Are More Than 20 test")
	@Test
	public void testCheckAllLongsAreMoreThan20() {
		assertAll(() -> assertFalse(jack.checkAllLongsAreMoreThan20(new int[] { 56, 13, 22, 113, 12, 6 })),
				() -> assertTrue(jack.checkAllLongsAreMoreThan20(new int[] { 111, 2222, 622 })),
				() -> assertFalse(jack.checkAllLongsAreMoreThan20(new int[] { 6, 3222, 1122333 })),
				() -> assertTrue(jack.checkAllLongsAreMoreThan20(new int[] { 21 })));

	}

	@DisplayName("Find Max Of All Numbers test")
	@Test
	public void testFindMaxOfAllNumbers() {
		assertAll(() -> assertEquals(113, jack.findMaxOfAllNumbers(new int[] { 56, 13, 22, 113, 12, 6 })),
				() -> assertEquals(2222, jack.findMaxOfAllNumbers(new int[] { 111, 2222, 622 })),
				() -> assertEquals(1122333, jack.findMaxOfAllNumbers(new int[] { 6, 3222, 1122333 })),
				() -> assertEquals(21, jack.findMaxOfAllNumbers(new int[] { 21 })));

	}

	@DisplayName("Compare Max Of 2 Arrays")
	@Test
	public void testCompareMaxOf2Arrays() {
		assertAll(
				() -> assertEquals(-1,jack.compareMaxOf2Arrays(new int[] { 56, 13, 22, 113, 12, 6 },new int[] { 56, 12, 22, 1113, 12, 6 })),
				() -> assertEquals(1, jack.compareMaxOf2Arrays(new int[] { 111, 2222, 622 }, new int[] { 2, 62 })),
				() -> assertEquals(0, jack.compareMaxOf2Arrays(new int[] { 6, 11, 123 }, new int[] { 61, 111, 123 })),
				() -> assertEquals(-1, jack.compareMaxOf2Arrays(new int[] { 21 }, new int[] { 22, 22 }))
				);

	}

	@DisplayName("Check String Contains Number")
	@Test
	public void testCheckStringContainsNumber() {
		assertAll(
				() ->assertTrue(jack.checkStringContainsNumber("ABC123"), "This String does not contains number"),
				() ->assertFalse(jack.checkStringContainsNumber("ABC"), "This String does not contains number"),
				() ->assertTrue(jack.checkStringContainsNumber("123"), "This String does not contains number")
				);

	}

	@DisplayName("Check String Array Contains Number")
	@Test
	public void testCheckStringArrayContainsNumber() {
		assertAll(
				() ->assertTrue(jack.checkStringContainsNumber("ABC123"), "This String does not contains number"),
				() ->assertFalse(jack.checkStringContainsNumber("ABC"), "This String does not contains number"),
				() ->assertTrue(jack.checkStringContainsNumber("123"), "This String does not contains number")
		);

	}

	@DisplayName("Get Number Out of String Array Test")
	@Test
	public void testGetNumberOutOfStringArray() {
		assertAll(
				() ->assertEquals(1232, jack.getNumberOutOfStringArray(new String[] { "ABC123", "AC2" })),
				() ->assertEquals(1232, jack.getNumberOutOfStringArray(new String[] { "ABC123", "AC2", "ACF" })),
				() ->assertEquals(0, jack.getNumberOutOfStringArray(new String[] { "ABC", "AC", "ACF" }))
		);

	}

	@DisplayName("Get Number Out of String Array Test Then get Sum Test")
	@Test
	public void testSumEachNumberContainedInsideStringArray() {
		assertAll(
				() ->assertEquals(223, jack.sumEachNumberContainedInsideStringArray(new String[] { "ABC123", "AC100" })),
				() ->assertEquals(125, jack.sumEachNumberContainedInsideStringArray(new String[] { "ABC123", "AC2", "ACF" })),
				() ->assertEquals(0, jack.sumEachNumberContainedInsideStringArray(new String[] { "ABC", "AC", "ACF" }))
		);
	}

	@DisplayName("Check whether a boolean exists in String Test")
	@Test
	public void testCheckIfBooleanWordExistsInString() {
		assertAll(
				() ->assertTrue(jack.checkIfBooleanWordExistsInString("TrueStatementFalse")),
				() ->assertFalse(jack.checkIfBooleanWordExistsInString("truStatemente")),
				() ->assertFalse(jack.checkIfBooleanWordExistsInString("FalStatemente")),
				() ->assertFalse(jack.checkIfBooleanWordExistsInString("ssss"))
				);

	}

	@AfterEach
	public void tearDown() {
		jack = null;
	}

	@AfterAll
	public static void cleanUp() {
		System.out.println("End of The Test");
	}

}